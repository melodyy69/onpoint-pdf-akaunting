

from __future__ import annotations

import json
import os
import re
import secrets
import shutil
from collections import OrderedDict
from datetime import datetime, timezone
from decimal import Decimal
from functools import wraps
from pathlib import Path
from typing import Any, Callable, TypeVar
from uuid import uuid4

import mysql.connector
from flask import (
    Flask,
    abort,
    current_app,
    flash,
    g,
    jsonify,
    redirect,
    render_template,
    request,
    send_file,
    session,
    url_for,
)
from werkzeug.security import check_password_hash, generate_password_hash
from werkzeug.middleware.proxy_fix import ProxyFix
from werkzeug.utils import secure_filename

# Значения по умолчанию для локального запуска. На хостинге одноимённые
# переменные окружения имеют приоритет, поэтому секреты не нужно менять в коде.
LOGIN_USERNAME = os.environ.get("LOGIN_USERNAME", "startech")
LOGIN_PASSWORD = os.environ.get("LOGIN_PASSWORD", "12345678")

MYSQL_HOST = os.environ.get("MYSQL_HOST", "127.0.0.1")
MYSQL_PORT = int(os.environ.get("MYSQL_PORT", "3306"))
MYSQL_USER = os.environ.get("MYSQL_USER", "root")
MYSQL_PASSWORD = os.environ.get("MYSQL_PASSWORD", "")
MYSQL_DATABASE = os.environ.get("MYSQL_DATABASE", "pdf_converter")

from main import (
    DEFAULT_ACCOUNT_NAME,
    DEFAULT_CURRENCY_CODE,
    DEFAULT_CURRENCY_RATE,
    StatementParseError,
    Transaction,
    category_match_text,
    category_memory_key,
    extract_transactions,
    learn_categories_from_workbook,
    verify_excel,
    write_excel,
)


MAX_UPLOAD_BYTES = 30 * 1024 * 1024
Function = TypeVar("Function", bound=Callable[..., Any])


def _env_flag(name: str, default: bool = False) -> bool:
    value = os.environ.get(name)
    if value is None:
        return default
    return value.strip().casefold() in {"1", "true", "yes", "on"}


def utc_now() -> str:
    return datetime.now(timezone.utc).isoformat(timespec="seconds")


def _get_or_create_secret(instance_path: Path) -> str:
    secret_path = instance_path / "secret_key"
    if secret_path.exists():
        secret = secret_path.read_text(encoding="utf-8").strip()
        if secret:
            return secret
    secret = secrets.token_hex(32)
    secret_path.write_text(secret + "\n", encoding="utf-8")
    secret_path.chmod(0o600)
    return secret


def create_app(
    test_config: dict[str, Any] | None = None,
    instance_path: Path | None = None,
) -> Flask:
    hosted_instance_path = os.environ.get("INSTANCE_PATH", "").strip()
    root = instance_path or (
        Path(hosted_instance_path).expanduser()
        if hosted_instance_path
        else Path(__file__).with_name("instance")
    )
    root.mkdir(parents=True, exist_ok=True)
    jobs_directory = os.environ.get("JOBS_DIRECTORY", "").strip() or str(
        root / "jobs"
    )

    app = Flask(
        __name__,
        instance_path=str(root),
        instance_relative_config=True,
    )
    app.config.from_mapping(
        SECRET_KEY=os.environ.get("SECRET_KEY") or _get_or_create_secret(root),
        JOBS_DIRECTORY=jobs_directory,
        LOGIN_USERNAME=LOGIN_USERNAME,
        LOGIN_PASSWORD=LOGIN_PASSWORD,
        MYSQL_HOST=MYSQL_HOST,
        MYSQL_PORT=MYSQL_PORT,
        MYSQL_USER=MYSQL_USER,
        MYSQL_PASSWORD=MYSQL_PASSWORD,
        MYSQL_DATABASE=MYSQL_DATABASE,
        MAX_CONTENT_LENGTH=MAX_UPLOAD_BYTES,
        CSRF_ENABLED=True,
        SESSION_COOKIE_HTTPONLY=True,
        SESSION_COOKIE_SAMESITE="Lax",
        SESSION_COOKIE_SECURE=_env_flag("SESSION_COOKIE_SECURE"),
    )
    if test_config:
        app.config.update(test_config)

    if _env_flag("TRUST_PROXY"):
        app.wsgi_app = ProxyFix(
            app.wsgi_app,
            x_for=1,
            x_proto=1,
            x_host=1,
        )

    Path(app.config["JOBS_DIRECTORY"]).mkdir(parents=True, exist_ok=True)
    app.teardown_appcontext(close_db)

    try:
        with app.app_context():
            ensure_mysql_database()
            init_database()
            user_id = ensure_configured_user()
            import_existing_files(user_id)
    except mysql.connector.Error as exc:
        raise RuntimeError(
            "Не удалось подключиться к MySQL. Проверьте MYSQL_HOST, MYSQL_PORT, "
            "MYSQL_USER, MYSQL_PASSWORD и запущен ли сервер MySQL."
        ) from exc

    @app.before_request
    def load_logged_in_user() -> None:
        user_id = session.get("user_id")
        g.user = (
            get_db()
            .execute("SELECT id, username FROM users WHERE id = %s", (user_id,))
            .fetchone()
            if user_id is not None
            else None
        )

    @app.before_request
    def validate_csrf() -> Any:
        if not app.config.get("CSRF_ENABLED", True):
            return None
        if request.method not in {"POST", "PUT", "PATCH", "DELETE"}:
            return None
        supplied = request.headers.get("X-CSRF-Token") or request.form.get("csrf_token")
        expected = session.get("csrf_token")
        if not supplied or not expected or not secrets.compare_digest(supplied, expected):
            if request.path.startswith("/api/"):
                return jsonify({"error": "Сессия устарела. Обновите страницу."}), 400
            abort(400, "Сессия устарела. Обновите страницу.")
        return None

    @app.context_processor
    def inject_template_helpers() -> dict[str, Any]:
        return {"csrf_token": get_csrf_token}

    @app.after_request
    def add_security_headers(response: Any) -> Any:
        response.headers["X-Content-Type-Options"] = "nosniff"
        response.headers["X-Frame-Options"] = "DENY"
        response.headers["Referrer-Policy"] = "no-referrer"
        response.headers["Content-Security-Policy"] = (
            "default-src 'self'; style-src 'self'; script-src 'self'; "
            "img-src 'self' data:; object-src 'none'; base-uri 'self'; "
            "form-action 'self'"
        )
        if g.get("user") is not None:
            response.headers["Cache-Control"] = "no-store"
        return response

    @app.route("/health")
    def health() -> Any:
        return {"status": "ok"}

    @app.get("/setup")
    def setup() -> Any:
        return redirect(url_for("login"))

    @app.route("/login", methods=("GET", "POST"))
    def login() -> Any:
        if g.user is not None:
            return redirect(url_for("index"))
        if request.method == "POST":
            username = request.form.get("username", "").strip().casefold()
            password = request.form.get("password", "")
            expected_username, expected_password = configured_credentials()
            credentials_match = secrets.compare_digest(
                username, expected_username
            ) and secrets.compare_digest(password, expected_password)
            user = get_db().execute(
                "SELECT id, username FROM users WHERE username = %s",
                (expected_username,),
            ).fetchone()
            if not credentials_match or user is None:
                flash("Неверный логин или пароль.", "error")
            else:
                session.clear()
                session["user_id"] = user["id"]
                flash("Вы вошли в систему.", "success")
                return redirect(url_for("index"))
        return render_template("login.html")

    @app.post("/logout")
    @login_required
    def logout() -> Any:
        session.clear()
        return redirect(url_for("login"))

    @app.route("/")
    @login_required
    def index() -> Any:
        rule_count = get_db().execute(
            "SELECT COUNT(*) AS count FROM category_rules WHERE user_id = %s",
            (g.user["id"],),
        ).fetchone()["count"]
        recent_jobs = get_db().execute(
            "SELECT id, original_filename, created_at FROM jobs "
            "WHERE user_id = %s ORDER BY created_at DESC LIMIT 5",
            (g.user["id"],),
        ).fetchall()
        return render_template(
            "index.html",
            rule_count=rule_count,
            recent_jobs=recent_jobs,
            account_name=DEFAULT_ACCOUNT_NAME,
            currency_code=DEFAULT_CURRENCY_CODE,
        )

    @app.post("/upload")
    @login_required
    def upload() -> Any:
        uploaded = request.files.get("pdf")
        if uploaded is None or not uploaded.filename:
            flash("Выберите банковскую выписку PDF.", "error")
            return redirect(url_for("index"))

        original_filename = Path(uploaded.filename).name
        safe_filename = secure_filename(original_filename) or "statement.pdf"
        if Path(safe_filename).suffix.casefold() != ".pdf":
            flash("Можно загрузить только файл PDF.", "error")
            return redirect(url_for("index"))

        job_id = uuid4().hex
        job_directory = Path(current_app.config["JOBS_DIRECTORY"]) / job_id
        job_directory.mkdir(parents=True, exist_ok=False)
        pdf_path = job_directory / "statement.pdf"
        uploaded.save(pdf_path)

        try:
            if pdf_path.read_bytes()[:5] != b"%PDF-":
                raise StatementParseError("Загруженный файл не является PDF.")
            transactions = extract_transactions(pdf_path)
            category_memory = get_category_memory(g.user["id"])
            output_name = f"{Path(safe_filename).stem}_akaunting.xlsx"
            xlsx_path = job_directory / output_name
            write_excel(
                transactions,
                xlsx_path,
                category_memory=category_memory,
                overwrite=True,
            )
            verify_excel(
                xlsx_path,
                len(transactions),
                category_memory,
                DEFAULT_ACCOUNT_NAME,
                DEFAULT_CURRENCY_CODE,
                DEFAULT_CURRENCY_RATE,
            )
        except (FileNotFoundError, StatementParseError, ValueError) as exc:
            shutil.rmtree(job_directory, ignore_errors=True)
            flash(f"Не удалось обработать выписку: {exc}", "error")
            return redirect(url_for("index"))

        timestamp = utc_now()
        get_db().execute(
            "INSERT INTO jobs "
            "(id, user_id, original_filename, pdf_path, xlsx_path, "
            "transactions_json, created_at, updated_at) "
            "VALUES (%s, %s, %s, %s, %s, %s, %s, %s)",
            (
                job_id,
                g.user["id"],
                original_filename,
                str(pdf_path),
                str(xlsx_path),
                serialize_transactions(transactions),
                timestamp,
                timestamp,
            ),
        )
        get_db().commit()
        return redirect(url_for("review_job", job_id=job_id))

    @app.route("/jobs/<job_id>")
    @login_required
    def review_job(job_id: str) -> Any:
        job = get_job(job_id, g.user["id"])
        transactions = deserialize_transactions(job["transactions_json"])
        category_memory = get_category_memory(g.user["id"])
        unknown_groups = group_unknown_transactions(transactions, category_memory)
        known_rows = sum(
            category_memory_key(item.transaction_type, item.description)
            in category_memory
            for item in transactions
        )
        category_suggestions = [
            row["category_name"]
            for row in get_db().execute(
                "SELECT DISTINCT category_name FROM category_rules "
                "WHERE user_id = %s "
                "ORDER BY category_name COLLATE utf8mb4_unicode_ci",
                (g.user["id"],),
            ).fetchall()
        ]
        return render_template(
            "review.html",
            job=job,
            transaction_count=len(transactions),
            known_rows=known_rows,
            unknown_groups=unknown_groups,
            category_suggestions=category_suggestions,
        )

    @app.post("/api/jobs/<job_id>/category")
    @login_required
    def save_job_category(job_id: str) -> Any:
        job = get_job(job_id, g.user["id"])
        payload = request.get_json(silent=True) or {}
        rule_key = str(payload.get("rule_key", "")).strip()
        category_name = str(payload.get("category_name", "")).strip()
        if not category_name:
            return jsonify({"error": "Введите название категории."}), 400
        if len(category_name) > 120:
            return jsonify({"error": "Название категории слишком длинное."}), 400

        transactions = deserialize_transactions(job["transactions_json"])
        matching = next(
            (
                item
                for item in transactions
                if category_memory_key(item.transaction_type, item.description)
                == rule_key
            ),
            None,
        )
        if matching is None:
            return jsonify({"error": "Транзакция не найдена в этой выписке."}), 400

        kind, description_key = rule_key.split("|", 1)
        timestamp = utc_now()
        get_db().execute(
            "INSERT INTO category_rules "
            "(user_id, transaction_type, description_key, example_description, "
            "category_name, created_at, updated_at) "
            "VALUES (%s, %s, %s, %s, %s, %s, %s) "
            "ON DUPLICATE KEY UPDATE "
            "example_description = VALUES(example_description), "
            "category_name = VALUES(category_name), "
            "updated_at = VALUES(updated_at)",
            (
                g.user["id"],
                kind,
                description_key,
                category_match_text(matching.description),
                category_name,
                timestamp,
                timestamp,
            ),
        )
        get_db().execute(
            "UPDATE jobs SET updated_at = %s WHERE id = %s AND user_id = %s",
            (timestamp, job_id, g.user["id"]),
        )
        get_db().commit()

        category_memory = get_category_memory(g.user["id"])
        xlsx_path = Path(job["xlsx_path"])
        write_excel(
            transactions,
            xlsx_path,
            category_memory=category_memory,
            overwrite=True,
        )
        verify_excel(
            xlsx_path,
            len(transactions),
            category_memory,
            DEFAULT_ACCOUNT_NAME,
            DEFAULT_CURRENCY_CODE,
            DEFAULT_CURRENCY_RATE,
        )
        remaining = group_unknown_transactions(transactions, category_memory)
        affected_rows = sum(
            category_memory_key(item.transaction_type, item.description) == rule_key
            for item in transactions
        )
        return jsonify(
            {
                "saved": True,
                "category_name": category_name,
                "affected_rows": affected_rows,
                "remaining_groups": len(remaining),
            }
        )

    @app.route("/jobs/<job_id>/download")
    @login_required
    def download_job(job_id: str) -> Any:
        job = get_job(job_id, g.user["id"])
        xlsx_path = Path(job["xlsx_path"])
        if not xlsx_path.is_file():
            abort(404)
        return send_file(
            xlsx_path,
            as_attachment=True,
            download_name=xlsx_path.name,
            mimetype=(
                "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            ),
        )

    @app.errorhandler(413)
    def upload_too_large(_error: Any) -> tuple[Any, int]:
        return render_template(
            "error.html",
            title="Файл слишком большой",
            message="Максимальный размер PDF — 30 МБ.",
        ), 413

    @app.errorhandler(400)
    def bad_request(error: Any) -> tuple[Any, int]:
        return render_template(
            "error.html",
            title="Не удалось выполнить запрос",
            message=getattr(error, "description", "Проверьте данные и повторите."),
        ), 400

    return app


class DatabaseResult:
    """Small result wrapper that closes MySQL cursors after fetching rows."""

    def __init__(self, cursor: Any) -> None:
        self._cursor = cursor if cursor.with_rows else None
        self.lastrowid = cursor.lastrowid
        self.rowcount = cursor.rowcount
        if self._cursor is None:
            cursor.close()

    def fetchone(self) -> dict[str, Any] | None:
        if self._cursor is None:
            return None
        try:
            return self._cursor.fetchone()
        finally:
            self._cursor.close()
            self._cursor = None

    def fetchall(self) -> list[dict[str, Any]]:
        if self._cursor is None:
            return []
        try:
            return self._cursor.fetchall()
        finally:
            self._cursor.close()
            self._cursor = None


class MySQLDatabase:
    """Expose the small database interface used by the Flask routes."""

    def __init__(self, connection: Any) -> None:
        self.connection = connection

    def execute(
        self, query: str, parameters: tuple[Any, ...] = ()
    ) -> DatabaseResult:
        cursor = self.connection.cursor(dictionary=True, buffered=True)
        cursor.execute(query, parameters)
        return DatabaseResult(cursor)

    def commit(self) -> None:
        self.connection.commit()

    def rollback(self) -> None:
        self.connection.rollback()

    def close(self) -> None:
        self.connection.close()


def mysql_connection_settings(include_database: bool = True) -> dict[str, Any]:
    """Return validated MySQL Connector/Python connection settings."""
    database_name = str(current_app.config["MYSQL_DATABASE"]).strip()
    if not re.fullmatch(r"[A-Za-z0-9_]+", database_name):
        raise RuntimeError(
            "MYSQL_DATABASE может содержать только латинские буквы, цифры и _."
        )
    settings: dict[str, Any] = {
        "host": str(current_app.config["MYSQL_HOST"]).strip(),
        "port": int(current_app.config["MYSQL_PORT"]),
        "user": str(current_app.config["MYSQL_USER"]),
        "password": str(current_app.config["MYSQL_PASSWORD"]),
        "charset": "utf8mb4",
        "collation": "utf8mb4_unicode_ci",
        "connection_timeout": 10,
        "use_pure": True,
        "autocommit": False,
    }
    if include_database:
        settings["database"] = database_name
    return settings


def ensure_mysql_database() -> None:
    """Use the configured database or create it when it does not exist."""
    settings = mysql_connection_settings(include_database=True)
    try:
        connection = mysql.connector.connect(**settings)
    except mysql.connector.Error as exc:
        if exc.errno != 1049:  # Unknown database
            raise
        server_settings = mysql_connection_settings(include_database=False)
        connection = mysql.connector.connect(**server_settings)
        cursor = connection.cursor()
        try:
            database_name = str(current_app.config["MYSQL_DATABASE"]).strip()
            cursor.execute(
                f"CREATE DATABASE IF NOT EXISTS `{database_name}` "
                "CHARACTER SET utf8mb4 COLLATE utf8mb4_unicode_ci"
            )
            connection.commit()
        finally:
            cursor.close()
            connection.close()
        return
    connection.close()


def get_db() -> MySQLDatabase:
    if "db" not in g:
        connection = mysql.connector.connect(**mysql_connection_settings())
        g.db = MySQLDatabase(connection)
    return g.db


def close_db(_error: BaseException | None = None) -> None:
    database = g.pop("db", None)
    if database is not None:
        database.close()


def init_database() -> None:
    database = get_db()
    schema_statements = (
        """
        CREATE TABLE IF NOT EXISTS users (
            id BIGINT UNSIGNED NOT NULL AUTO_INCREMENT,
            username VARCHAR(50) NOT NULL,
            password_hash VARCHAR(255) NOT NULL,
            created_at VARCHAR(40) NOT NULL,
            PRIMARY KEY (id),
            UNIQUE KEY uq_users_username (username)
        ) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_unicode_ci
        """,
        """
        CREATE TABLE IF NOT EXISTS category_rules (
            id BIGINT UNSIGNED NOT NULL AUTO_INCREMENT,
            user_id BIGINT UNSIGNED NOT NULL,
            transaction_type ENUM('income', 'expense') NOT NULL,
            description_key VARCHAR(255) NOT NULL,
            example_description TEXT NOT NULL,
            category_name VARCHAR(120) NOT NULL,
            created_at VARCHAR(40) NOT NULL,
            updated_at VARCHAR(40) NOT NULL,
            PRIMARY KEY (id),
            UNIQUE KEY uq_category_rule
                (user_id, transaction_type, description_key),
            CONSTRAINT fk_category_rules_user
                FOREIGN KEY (user_id) REFERENCES users(id) ON DELETE CASCADE
        ) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_unicode_ci
        """,
        """
        CREATE TABLE IF NOT EXISTS jobs (
            id CHAR(32) NOT NULL,
            user_id BIGINT UNSIGNED NOT NULL,
            original_filename VARCHAR(255) NOT NULL,
            pdf_path TEXT NOT NULL,
            xlsx_path TEXT NOT NULL,
            transactions_json LONGTEXT NOT NULL,
            created_at VARCHAR(40) NOT NULL,
            updated_at VARCHAR(40) NOT NULL,
            PRIMARY KEY (id),
            KEY idx_jobs_user_created_at (user_id, created_at),
            CONSTRAINT fk_jobs_user
                FOREIGN KEY (user_id) REFERENCES users(id) ON DELETE CASCADE
        ) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_unicode_ci
        """,
    )
    for statement in schema_statements:
        database.execute(statement)
    database.commit()


def configured_credentials() -> tuple[str, str]:
    """Return the single local account configured at the top of this file."""
    username = str(current_app.config["LOGIN_USERNAME"]).strip().casefold()
    password = str(current_app.config["LOGIN_PASSWORD"])
    if not username:
        raise RuntimeError("LOGIN_USERNAME не может быть пустым.")
    if not password:
        raise RuntimeError("LOGIN_PASSWORD не может быть пустым.")
    return username, password


def ensure_configured_user() -> int:
    """Create or update the internal user while preserving its related data."""
    username, password = configured_credentials()
    database = get_db()
    user = database.execute(
        "SELECT id, username, password_hash FROM users ORDER BY id LIMIT 1"
    ).fetchone()
    password_hash = generate_password_hash(password)
    if user is None:
        result = database.execute(
            "INSERT INTO users (username, password_hash, created_at) "
            "VALUES (%s, %s, %s)",
            (username, password_hash, utc_now()),
        )
        user_id = int(result.lastrowid)
    elif user["username"] != username or not check_password_hash(
        user["password_hash"], password
    ):
        database.execute(
            "UPDATE users SET username = %s, password_hash = %s WHERE id = %s",
            (username, password_hash, user["id"]),
        )
        user_id = int(user["id"])
    else:
        user_id = int(user["id"])
    database.commit()
    return user_id


def import_existing_files(user_id: int) -> None:
    """Restore existing local workbooks and jobs into an empty/new MySQL DB."""
    jobs_root = Path(current_app.config["JOBS_DIRECTORY"])
    database = get_db()
    for job_directory in sorted(jobs_root.iterdir()):
        if not job_directory.is_dir():
            continue
        pdf_path = job_directory / "statement.pdf"
        workbooks = sorted(job_directory.glob("*.xlsx"))
        if not pdf_path.is_file() or not workbooks:
            continue
        xlsx_path = workbooks[0]
        timestamp = datetime.fromtimestamp(
            xlsx_path.stat().st_mtime, timezone.utc
        ).isoformat(timespec="seconds")

        category_memory: dict[str, str] = {}
        try:
            learn_categories_from_workbook(
                xlsx_path,
                category_memory,
                require_marker=True,
            )
        except (FileNotFoundError, StatementParseError, OSError):
            category_memory = {}
        for rule_key, category_name in category_memory.items():
            kind, description_key = rule_key.split("|", 1)
            database.execute(
                "INSERT IGNORE INTO category_rules "
                "(user_id, transaction_type, description_key, "
                "example_description, category_name, created_at, updated_at) "
                "VALUES (%s, %s, %s, %s, %s, %s, %s)",
                (
                    user_id,
                    kind,
                    description_key,
                    description_key,
                    category_name,
                    timestamp,
                    timestamp,
                ),
            )

        existing_job = database.execute(
            "SELECT id FROM jobs WHERE id = %s",
            (job_directory.name,),
        ).fetchone()
        if existing_job is not None:
            continue
        try:
            transactions = extract_transactions(pdf_path)
        except (FileNotFoundError, StatementParseError, ValueError):
            continue
        output_stem = xlsx_path.stem
        if output_stem.endswith("_akaunting"):
            output_stem = output_stem[: -len("_akaunting")]
        database.execute(
            "INSERT IGNORE INTO jobs "
            "(id, user_id, original_filename, pdf_path, xlsx_path, "
            "transactions_json, created_at, updated_at) "
            "VALUES (%s, %s, %s, %s, %s, %s, %s, %s)",
            (
                job_directory.name,
                user_id,
                f"{output_stem}.pdf",
                str(pdf_path),
                str(xlsx_path),
                serialize_transactions(transactions),
                timestamp,
                timestamp,
            ),
        )
    database.commit()


def get_csrf_token() -> str:
    token = session.get("csrf_token")
    if token is None:
        token = secrets.token_urlsafe(32)
        session["csrf_token"] = token
    return token


def login_required(view: Function) -> Function:
    @wraps(view)
    def wrapped_view(*args: Any, **kwargs: Any) -> Any:
        if g.user is None:
            return redirect(url_for("login"))
        return view(*args, **kwargs)

    return wrapped_view  # type: ignore[return-value]


def get_job(job_id: str, user_id: int) -> dict[str, Any]:
    job = get_db().execute(
        "SELECT * FROM jobs WHERE id = %s AND user_id = %s",
        (job_id, user_id),
    ).fetchone()
    if job is None:
        abort(404)
    return job


def get_category_memory(user_id: int) -> dict[str, str]:
    rows = get_db().execute(
        "SELECT transaction_type, description_key, category_name "
        "FROM category_rules WHERE user_id = %s",
        (user_id,),
    ).fetchall()
    return {
        f"{row['transaction_type']}|{row['description_key']}": row["category_name"]
        for row in rows
    }


def serialize_transactions(transactions: list[Transaction]) -> str:
    return json.dumps(
        [
            {
                "paid_at": item.paid_at.isoformat(),
                "signed_amount": str(item.signed_amount),
                "balance": str(item.balance),
                "description": item.description,
                "page_number": item.page_number,
            }
            for item in transactions
        ],
        ensure_ascii=False,
        separators=(",", ":"),
    )


def deserialize_transactions(payload: str) -> list[Transaction]:
    return [
        Transaction(
            paid_at=datetime.fromisoformat(item["paid_at"]),
            signed_amount=Decimal(item["signed_amount"]),
            balance=Decimal(item["balance"]),
            description=item["description"],
            page_number=int(item["page_number"]),
        )
        for item in json.loads(payload)
    ]


def group_unknown_transactions(
    transactions: list[Transaction], category_memory: dict[str, str]
) -> list[dict[str, Any]]:
    groups: OrderedDict[str, dict[str, Any]] = OrderedDict()
    for item in transactions:
        kind = item.transaction_type.lower()
        key = category_memory_key(kind, item.description)
        if key in category_memory:
            continue
        group = groups.setdefault(
            key,
            {
                "rule_key": key,
                "transaction_type": kind,
                "merchant_text": category_match_text(item.description),
                "description": item.description,
                "count": 0,
                "dates": [],
                "amounts": [],
            },
        )
        group["count"] += 1
        date_value = item.paid_at.strftime("%Y-%m-%d")
        amount_value = f"{item.amount:.2f}"
        if date_value not in group["dates"] and len(group["dates"]) < 3:
            group["dates"].append(date_value)
        if amount_value not in group["amounts"] and len(group["amounts"]) < 3:
            group["amounts"].append(amount_value)
    return list(groups.values())


if __name__ == "__main__":
    app = create_app()
    app.run(
        host=os.environ.get("HOST", "0.0.0.0"),
        port=int(os.environ.get("PORT", "5000")),
        debug=False,
    )
