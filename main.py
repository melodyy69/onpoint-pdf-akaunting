"""Convert an OnPoint bank statement PDF to an Akaunting Excel import file.

The statement is read with pdfplumber from page 2 through the final page.
The output follows the 12-column Akaunting transaction import template.
"""

from __future__ import annotations

import argparse
import json
import re
import sys
import unicodedata
from dataclasses import dataclass
from datetime import datetime
from decimal import Decimal, InvalidOperation
from pathlib import Path
from typing import Any, Iterable

import pdfplumber
from openpyxl import Workbook, load_workbook
from openpyxl.utils import get_column_letter


DEFAULT_ACCOUNT_NAME = "Business Account"
DEFAULT_CURRENCY_CODE = "USD"
DEFAULT_CURRENCY_RATE = Decimal("1")
DEFAULT_START_PAGE = 2
TRANSACTION_NUMBER_PREFIX = "TRA-"
DEFAULT_PAYMENT_METHOD = "offline-payments.cash.1"
CATEGORY_MEMORY_FILENAME = "category_memory.json"
CATEGORY_MEMORY_VERSION = 2
CATEGORY_MEMORY_WORKBOOK_MARKER = "pdf_conventer_category_memory_v1"

HEADERS = [
    "type",
    "number",
    "paid_at",
    "amount",
    "currency_code",
    "currency_rate",
    "account_name",
    "invoice_bill_number",
    "contact_email",
    "category_name",
    "description",
    "payment_method",
]

DATE_RE = re.compile(r"^\d{2}/\d{2}/\d{4}$")
MONEY_RE = re.compile(r"^[+-]?\$?\(?\d[\d,]*\.\d{2}\)?$")
SPACE_RE = re.compile(r"\s+")
POS_PREFIX_RE = re.compile(
    r"^point\s+of\s+sale\s+(?:withdrawal|deposit)\b\s*",
    re.IGNORECASE,
)
MERCHANT_TOKEN_RE = re.compile(r"[^\W_]+(?:[-'][^\W_]+)*", re.UNICODE)
MERCHANT_TOKEN_LIMIT = 2
PROCESSOR_PREFIXES = {"py", "sq", "sp", "tst"}

INCOME_WORDS = ("deposit", "credit", "refund", "interest", "reversal")
EXPENSE_WORDS = (
    "withdrawal",
    "debit",
    "fee",
    "payment",
    "purchase",
    "check",
)


class StatementParseError(RuntimeError):
    """Raised when a statement cannot be parsed or reconciled safely."""


@dataclass(frozen=True)
class Transaction:
    paid_at: datetime
    signed_amount: Decimal
    balance: Decimal
    description: str
    page_number: int

    @property
    def amount(self) -> Decimal:
        return abs(self.signed_amount)

    @property
    def transaction_type(self) -> str:
        return "Income" if self.signed_amount >= 0 else "Expense"


@dataclass(frozen=True)
class StatementSummary:
    income_count: int
    income_total: Decimal
    expense_count: int
    expense_total: Decimal


def normalize_money_text(value: str) -> str:
    """Normalize PDF minus/currency characters without changing the value."""
    return (
        value.strip()
        .replace("−", "-")
        .replace("–", "-")
        .replace("—", "-")
        .replace(" ", "")
    )


def parse_money(value: str) -> Decimal:
    """Parse a bank-formatted amount into an exact Decimal."""
    normalized = normalize_money_text(value)
    negative_parentheses = normalized.startswith("(") and normalized.endswith(")")
    normalized = normalized.replace("$", "").replace(",", "")
    if negative_parentheses:
        normalized = f"-{normalized[1:-1]}"
    try:
        return Decimal(normalized)
    except InvalidOperation as exc:
        raise StatementParseError(f"Некорректная сумма в PDF: {value!r}") from exc


def normalize_description(value: str) -> str:
    """Normalize a description for deterministic category matching."""
    normalized = unicodedata.normalize("NFKC", value)
    return SPACE_RE.sub(" ", normalized.strip()).casefold()


def pos_merchant_text(description: str) -> str:
    """Return the text after a Point Of Sale Withdrawal/Deposit prefix."""
    normalized = unicodedata.normalize("NFKC", description)
    normalized = SPACE_RE.sub(" ", normalized.strip())
    match = POS_PREFIX_RE.match(normalized)
    if match is None:
        return ""
    return normalized[match.end() :].strip(" -*:#")


def category_match_text(description: str) -> str:
    """Return the user-facing merchant text used for a category rule."""
    return pos_merchant_text(description) or SPACE_RE.sub(
        " ", unicodedata.normalize("NFKC", description).strip()
    )


def merchant_match_key(description: str) -> str:
    """Build a stable merchant key for POS descriptions.

    Bank descriptions often append a branch number, authorization code, phone,
    address, or city to the same merchant. For POS rows, the bank prefix is
    removed and only the first stable merchant words are kept. Non-POS rows
    retain exact normalized matching.
    """
    merchant_text = pos_merchant_text(description)
    if not merchant_text:
        return normalize_description(description)

    tokens = MERCHANT_TOKEN_RE.findall(normalize_description(merchant_text))
    if tokens and tokens[0] in PROCESSOR_PREFIXES:
        tokens = tokens[1:]

    stable_tokens: list[str] = []
    for token in tokens:
        has_digit = any(character.isdigit() for character in token)
        if has_digit and stable_tokens:
            break
        stable_tokens.append(token)
        if len(stable_tokens) >= MERCHANT_TOKEN_LIMIT:
            break

    return " ".join(stable_tokens) or normalize_description(merchant_text)


def category_memory_key(transaction_type: str, description: str) -> str:
    """Build the persistent key used to remember a manual category."""
    kind = transaction_type.strip().casefold()
    merchant_key = merchant_match_key(description)
    if kind not in {"income", "expense"} or not merchant_key:
        return ""
    return f"{kind}|{merchant_key}"


def load_category_memory(memory_path: Path) -> dict[str, str]:
    """Load remembered categories from JSON, or return an empty memory."""
    if not memory_path.exists():
        return {}
    try:
        payload = json.loads(memory_path.read_text(encoding="utf-8"))
    except (OSError, UnicodeError, json.JSONDecodeError) as exc:
        raise StatementParseError(
            f"Не удалось прочитать память категорий: {memory_path}"
        ) from exc

    if not isinstance(payload, dict) or payload.get("version") not in {1, 2}:
        raise StatementParseError(
            f"Неподдерживаемый формат памяти категорий: {memory_path}"
        )
    categories = payload.get("categories")
    if not isinstance(categories, dict):
        raise StatementParseError(
            f"В памяти категорий отсутствует объект categories: {memory_path}"
        )

    memory: dict[str, str] = {}
    for key, category in categories.items():
        if not isinstance(key, str) or not isinstance(category, str):
            continue
        if not category.strip() or "|" not in key:
            continue
        kind, stored_description = key.split("|", 1)
        migrated_key = category_memory_key(kind, stored_description)
        if migrated_key:
            memory[migrated_key] = category.strip()
    return memory


def save_category_memory(memory_path: Path, memory: dict[str, str]) -> None:
    """Save remembered categories atomically."""
    memory_path.parent.mkdir(parents=True, exist_ok=True)
    temporary_path = memory_path.with_name(f".{memory_path.name}.tmp")
    payload = {
        "version": CATEGORY_MEMORY_VERSION,
        "categories": dict(sorted(memory.items())),
    }
    try:
        temporary_path.write_text(
            json.dumps(payload, ensure_ascii=False, indent=2) + "\n",
            encoding="utf-8",
        )
        temporary_path.replace(memory_path)
    finally:
        if temporary_path.exists():
            temporary_path.unlink()


def learn_categories_from_workbook(
    workbook_path: Path,
    memory: dict[str, str],
    require_marker: bool,
) -> int:
    """Learn non-empty manual categories from an Akaunting workbook."""
    if not workbook_path.is_file():
        raise FileNotFoundError(f"Excel для обучения не найден: {workbook_path}")
    try:
        workbook = load_workbook(workbook_path, data_only=True, read_only=True)
    except Exception as exc:
        raise StatementParseError(
            f"Не удалось прочитать Excel для обучения: {workbook_path}"
        ) from exc

    try:
        keywords = workbook.properties.keywords or ""
        if require_marker and CATEGORY_MEMORY_WORKBOOK_MARKER not in keywords:
            return 0
        if "Transactions" not in workbook.sheetnames:
            if require_marker:
                return 0
            raise StatementParseError(
                f"В Excel для обучения нет листа Transactions: {workbook_path}"
            )

        sheet = workbook["Transactions"]
        headers = [cell.value for cell in sheet[1]]
        columns = {
            str(value).strip(): index
            for index, value in enumerate(headers)
            if value is not None
        }
        required = {"type", "category_name", "description"}
        if not required.issubset(columns):
            if require_marker:
                return 0
            raise StatementParseError(
                f"В Excel для обучения нет колонок type, category_name и description: "
                f"{workbook_path}"
            )

        changes = 0
        for row in sheet.iter_rows(min_row=2, values_only=True):
            kind = row[columns["type"]]
            category = row[columns["category_name"]]
            description = row[columns["description"]]
            if not all(isinstance(value, str) for value in (kind, category, description)):
                continue
            category = category.strip()
            key = category_memory_key(kind, description)
            if not category or not key:
                continue
            if memory.get(key) != category:
                memory[key] = category
                changes += 1
        return changes
    finally:
        workbook.close()


def update_category_memory(
    memory_path: Path,
    output_directory: Path,
    explicit_workbooks: Iterable[Path],
    auto_learn: bool,
) -> tuple[dict[str, str], int]:
    """Load memory and update it from manually categorized workbooks."""
    memory = load_category_memory(memory_path)
    changes = 0
    seen: set[Path] = set()

    if auto_learn and output_directory.is_dir():
        candidates = sorted(
            (
                path
                for path in output_directory.glob("*.xlsx")
                if not path.name.startswith("~$")
            ),
            key=lambda path: (path.stat().st_mtime, path.name.casefold()),
        )
        for workbook_path in candidates:
            resolved = workbook_path.resolve()
            seen.add(resolved)
            try:
                changes += learn_categories_from_workbook(
                    resolved, memory, require_marker=True
                )
            except (FileNotFoundError, StatementParseError):
                continue

    for workbook_path in explicit_workbooks:
        resolved = workbook_path.expanduser().resolve()
        if resolved not in seen:
            seen.add(resolved)
        changes += learn_categories_from_workbook(
            resolved, memory, require_marker=False
        )

    if changes:
        save_category_memory(memory_path, memory)
    return memory, changes


def _same_line(
    words: Iterable[dict[str, Any]], top: float, tolerance: float = 2.0
) -> list[dict[str, Any]]:
    return sorted(
        (word for word in words if abs(float(word["top"]) - top) <= tolerance),
        key=lambda word: float(word["x0"]),
    )


def _column_boundaries(page: Any, words: list[dict[str, Any]]) -> tuple[float, float]:
    """Return left edges for the Amount and Balance columns."""
    amount_headers = sorted(
        (word for word in words if word["text"].strip().lower() == "amount"),
        key=lambda word: float(word["top"]),
    )
    balance_headers = [
        word for word in words if word["text"].strip().lower() == "balance"
    ]
    amount_header = amount_headers[0] if amount_headers else None
    matching_balance_headers = (
        sorted(
            (
                word
                for word in balance_headers
                if float(word["x0"]) > float(amount_header["x0"])
                and abs(float(word["top"]) - float(amount_header["top"])) <= 3.0
            ),
            key=lambda word: float(word["x0"]),
        )
        if amount_header
        else []
    )
    amount_left = (
        float(amount_header["x0"]) - 35.0
        if amount_header
        else float(page.width) * 0.69
    )
    balance_left = (
        float(matching_balance_headers[0]["x0"]) - 15.0
        if matching_balance_headers
        else float(page.width) * 0.86
    )
    if amount_left >= balance_left:
        raise StatementParseError("Не удалось определить колонки Amount и Balance.")
    return amount_left, balance_left


def _infer_signed_amount(raw_amount: str, description: str) -> Decimal:
    amount = parse_money(raw_amount)
    normalized = normalize_money_text(raw_amount)
    if amount < 0 or normalized.startswith("-") or normalized.startswith("("):
        return -abs(amount)

    description_lower = description.lower()
    if any(word in description_lower for word in INCOME_WORDS):
        return abs(amount)
    if any(word in description_lower for word in EXPENSE_WORDS):
        return -abs(amount)
    return abs(amount)


def _parse_page(page: Any, page_number: int) -> list[Transaction]:
    words = page.extract_words(x_tolerance=2, y_tolerance=3, keep_blank_chars=False)
    if not words:
        return []

    date_words = sorted(
        (
            word
            for word in words
            if float(word["x0"]) < float(page.width) * 0.13
            and DATE_RE.fullmatch(word["text"].strip())
        ),
        key=lambda word: float(word["top"]),
    )
    if not date_words:
        return []
    amount_left, balance_left = _column_boundaries(page, words)

    transactions: list[Transaction] = []
    for date_word in date_words:
        top = float(date_word["top"])
        row_words = _same_line(words, top)
        amount_words = [
            word
            for word in row_words
            if amount_left <= float(word["x0"]) < balance_left
            and MONEY_RE.fullmatch(normalize_money_text(word["text"]))
        ]
        balance_words = [
            word
            for word in row_words
            if float(word["x0"]) >= balance_left
            and MONEY_RE.fullmatch(normalize_money_text(word["text"]))
        ]

        if len(amount_words) != 1 or len(balance_words) != 1:
            raise StatementParseError(
                f"Страница {page_number}: для строки {date_word['text']} найдено "
                f"Amount={len(amount_words)}, Balance={len(balance_words)}."
            )

        description_words = [
            word
            for word in row_words
            if float(word["x0"]) > float(date_word["x1"]) + 5
            and float(word["x0"]) < amount_left
        ]
        description = SPACE_RE.sub(
            " ", " ".join(word["text"].strip() for word in description_words)
        ).strip()

        try:
            paid_at = datetime.strptime(date_word["text"], "%m/%d/%Y")
        except ValueError as exc:
            raise StatementParseError(
                f"Страница {page_number}: некорректная дата {date_word['text']!r}."
            ) from exc

        raw_amount = amount_words[0]["text"]
        transactions.append(
            Transaction(
                paid_at=paid_at,
                signed_amount=_infer_signed_amount(raw_amount, description),
                balance=parse_money(balance_words[0]["text"]),
                description=description,
                page_number=page_number,
            )
        )
    return transactions


def _extract_statement_summary(text: str) -> StatementSummary | None:
    credit_match = re.search(
        r"(?P<count>\d+)\s+Total Credits\s+(?P<amount>[\d,]+\.\d{2})",
        text,
        flags=re.IGNORECASE,
    )
    debit_match = re.search(
        r"(?P<count>\d+)\s+Total Debits\s+(?P<amount>-?[\d,]+\.\d{2})",
        text,
        flags=re.IGNORECASE,
    )
    if not credit_match or not debit_match:
        return None
    return StatementSummary(
        income_count=int(credit_match.group("count")),
        income_total=abs(parse_money(credit_match.group("amount"))),
        expense_count=int(debit_match.group("count")),
        expense_total=abs(parse_money(debit_match.group("amount"))),
    )


def _validate_balance_sequence(transactions: list[Transaction]) -> None:
    for previous, current in zip(transactions, transactions[1:]):
        expected_balance = previous.balance + current.signed_amount
        if abs(expected_balance - current.balance) > Decimal("0.01"):
            raise StatementParseError(
                "Нарушена сверка баланса: "
                f"после {current.paid_at:%Y-%m-%d} на странице {current.page_number} "
                f"ожидалось {expected_balance:.2f}, в PDF указано {current.balance:.2f}."
            )


def _validate_summary(
    transactions: list[Transaction], summary: StatementSummary
) -> None:
    incomes = [item for item in transactions if item.signed_amount >= 0]
    expenses = [item for item in transactions if item.signed_amount < 0]
    income_total = sum((item.amount for item in incomes), Decimal("0"))
    expense_total = sum((item.amount for item in expenses), Decimal("0"))

    actual = (len(incomes), income_total, len(expenses), expense_total)
    expected = (
        summary.income_count,
        summary.income_total,
        summary.expense_count,
        summary.expense_total,
    )
    if actual != expected:
        raise StatementParseError(
            "Операции не совпали с итогами выписки. "
            f"Извлечено Income: {actual[0]} / {actual[1]:.2f}, "
            f"Expense: {actual[2]} / {actual[3]:.2f}; "
            f"ожидалось Income: {expected[0]} / {expected[1]:.2f}, "
            f"Expense: {expected[2]} / {expected[3]:.2f}."
        )


def extract_transactions(
    pdf_path: Path, start_page: int = DEFAULT_START_PAGE, validate: bool = True
) -> list[Transaction]:
    """Extract transactions from start_page through the last PDF page."""
    if start_page < 1:
        raise ValueError("Номер первой страницы должен быть не меньше 1.")
    if not pdf_path.is_file():
        raise FileNotFoundError(f"PDF не найден: {pdf_path}")

    with pdfplumber.open(pdf_path) as pdf:
        if start_page > len(pdf.pages):
            raise StatementParseError(
                f"В PDF {len(pdf.pages)} страниц, начать со страницы {start_page} нельзя."
            )

        transactions: list[Transaction] = []
        summary: StatementSummary | None = None
        for page_number in range(start_page, len(pdf.pages) + 1):
            page = pdf.pages[page_number - 1]
            if summary is None:
                summary = _extract_statement_summary(page.extract_text() or "")
            transactions.extend(_parse_page(page, page_number))

    if not transactions:
        raise StatementParseError(
            f"Начиная со страницы {start_page}, операции в PDF не найдены."
        )
    if validate:
        _validate_balance_sequence(transactions)
        if summary is not None:
            _validate_summary(transactions, summary)
    return transactions


def write_excel(
    transactions: list[Transaction],
    output_path: Path,
    category_memory: dict[str, str],
    account_name: str = DEFAULT_ACCOUNT_NAME,
    currency_code: str = DEFAULT_CURRENCY_CODE,
    currency_rate: Decimal = DEFAULT_CURRENCY_RATE,
    overwrite: bool = False,
) -> int:
    """Write the Akaunting import workbook."""
    if not account_name:
        raise ValueError("Account Name не может быть пустым.")
    if not re.fullmatch(r"[A-Z]{3}", currency_code):
        raise ValueError("Currency Code должен состоять из трех заглавных букв.")
    if currency_rate <= 0:
        raise ValueError("Currency Rate должен быть больше нуля.")
    if output_path.exists() and not overwrite:
        raise FileExistsError(
            f"Файл уже существует: {output_path}. Добавьте --overwrite для замены."
        )

    output_path.parent.mkdir(parents=True, exist_ok=True)
    workbook = Workbook()
    workbook.properties.keywords = CATEGORY_MEMORY_WORKBOOK_MARKER
    sheet = workbook.active
    sheet.title = "Transactions"
    sheet.freeze_panes = "A2"

    sheet.append(HEADERS)
    remembered_count = 0
    for transaction_index, item in enumerate(transactions, start=1):
        kind = item.transaction_type.lower()
        remembered_category = category_memory.get(
            category_memory_key(kind, item.description)
        )
        if remembered_category:
            remembered_count += 1
        sheet.append(
            [
                kind,
                f"{TRANSACTION_NUMBER_PREFIX}{transaction_index:05d}",
                item.paid_at,
                float(item.amount),
                currency_code,
                float(currency_rate),
                account_name,
                None,
                None,
                remembered_category,
                item.description,
                DEFAULT_PAYMENT_METHOD,
            ]
        )

    for row in sheet.iter_rows(min_row=2, max_row=sheet.max_row):
        row[2].number_format = "yyyy-mm-dd"
        row[3].number_format = "0.00"
        row[5].number_format = (
            "0" if currency_rate == currency_rate.to_integral() else "0.########"
        )

    widths = [12, 16, 13, 16, 16, 15, 22, 22, 22, 20, 50, 28]
    for index, width in enumerate(widths, start=1):
        sheet.column_dimensions[get_column_letter(index)].width = width
    sheet.auto_filter.ref = f"A1:L{sheet.max_row}"

    temporary_path = output_path.with_name(f".{output_path.stem}.tmp.xlsx")
    try:
        workbook.save(temporary_path)
        temporary_path.replace(output_path)
    finally:
        if temporary_path.exists():
            temporary_path.unlink()
    return remembered_count


def verify_excel(
    output_path: Path,
    expected_rows: int,
    category_memory: dict[str, str],
    account_name: str,
    currency_code: str,
    currency_rate: Decimal,
) -> None:
    """Reopen the workbook and verify the import contract."""
    workbook = load_workbook(output_path, data_only=False, read_only=True)
    try:
        if workbook.sheetnames != ["Transactions"]:
            raise StatementParseError("В Excel должен быть один лист Transactions.")
        sheet = workbook["Transactions"]
        headers = [sheet.cell(row=1, column=column).value for column in range(1, 13)]
        if headers != HEADERS:
            raise StatementParseError(f"Некорректные заголовки Excel: {headers}")
        if sheet.max_row - 1 != expected_rows or sheet.max_column != 12:
            raise StatementParseError(
                f"Некорректный размер Excel: {sheet.max_row - 1} строк, "
                f"{sheet.max_column} колонок."
            )
        if sheet["C2"].number_format != "yyyy-mm-dd" or sheet["D2"].number_format != "0.00":
            raise StatementParseError("Некорректный формат paid_at или amount.")
        for transaction_index, row in enumerate(
            sheet.iter_rows(min_row=2, values_only=True), start=1
        ):
            (
                kind,
                transaction_number,
                paid_at,
                amount,
                code,
                rate,
                row_account,
                invoice_bill_number,
                contact_email,
                category,
                description,
                payment_method,
            ) = row
            if not isinstance(paid_at, datetime) or not paid_at.strftime("%Y-%m-%d"):
                raise StatementParseError("В paid_at обнаружена некорректная дата.")
            if not isinstance(amount, (int, float)) or amount < 0:
                raise StatementParseError("В amount обнаружена некорректная сумма.")
            if row_account != account_name or kind not in {"income", "expense"}:
                raise StatementParseError("Некорректные account_name или type.")
            if (
                transaction_number
                != f"{TRANSACTION_NUMBER_PREFIX}{transaction_index:05d}"
                or payment_method != DEFAULT_PAYMENT_METHOD
            ):
                raise StatementParseError(
                    "Некорректные number или payment_method."
                )
            expected_category = category_memory.get(
                category_memory_key(kind, description)
            )
            if (
                invoice_bill_number is not None
                or contact_email is not None
                or category != expected_category
                or not isinstance(description, str)
                or code != currency_code
                or Decimal(str(rate)) != currency_rate
            ):
                raise StatementParseError(
                    "Некорректные поля шаблона Akaunting."
                )
    finally:
        workbook.close()


def choose_pdf() -> Path:
    """Open a file picker when main.py is run without a PDF argument."""
    try:
        from tkinter import Tk, filedialog

        root = Tk()
        root.withdraw()
        root.attributes("-topmost", True)
        selected = filedialog.askopenfilename(
            title="Выберите банковскую выписку PDF",
            filetypes=[("PDF", "*.pdf")],
        )
        root.destroy()
    except Exception as exc:  # pragma: no cover - depends on desktop environment
        raise StatementParseError(
            "Не удалось открыть окно выбора файла. Передайте путь к PDF аргументом."
        ) from exc
    if not selected:
        raise StatementParseError("PDF не выбран.")
    return Path(selected)


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        description=(
            "Конвертирует банковскую PDF-выписку OnPoint в Excel для Akaunting. "
            "Чтение начинается со второй страницы."
        )
    )
    parser.add_argument("pdf", nargs="?", type=Path, help="Путь к банковскому PDF")
    parser.add_argument("-o", "--output", type=Path, help="Путь к итоговому .xlsx")
    parser.add_argument("--account-name", default=DEFAULT_ACCOUNT_NAME)
    parser.add_argument("--currency-code", default=DEFAULT_CURRENCY_CODE)
    parser.add_argument("--currency-rate", type=Decimal, default=DEFAULT_CURRENCY_RATE)
    parser.add_argument("--start-page", type=int, default=DEFAULT_START_PAGE)
    parser.add_argument(
        "--category-memory",
        type=Path,
        default=Path(__file__).with_name(CATEGORY_MEMORY_FILENAME),
        help="Путь к JSON-файлу памяти категорий",
    )
    parser.add_argument(
        "--learn-from",
        type=Path,
        action="append",
        default=[],
        help="Дополнительный вручную заполненный Excel для обучения (можно повторять)",
    )
    parser.add_argument(
        "--no-auto-learn",
        action="store_true",
        help="Не искать заполненные Excel в папке итогового файла",
    )
    parser.add_argument(
        "--skip-validation",
        action="store_true",
        help="Не сверять операции с балансами и итогами выписки",
    )
    parser.add_argument(
        "--overwrite", action="store_true", help="Разрешить замену существующего Excel"
    )
    return parser


def main(argv: list[str] | None = None) -> int:
    args = build_parser().parse_args(argv)
    try:
        pdf_path = args.pdf.expanduser().resolve() if args.pdf else choose_pdf().resolve()
        output_path = (
            args.output.expanduser().resolve()
            if args.output
            else pdf_path.with_name(f"{pdf_path.stem}_akaunting.xlsx")
        )
        transactions = extract_transactions(
            pdf_path,
            start_page=args.start_page,
            validate=not args.skip_validation,
        )
        memory_path = args.category_memory.expanduser().resolve()
        category_memory, learned_count = update_category_memory(
            memory_path,
            output_path.parent,
            args.learn_from,
            auto_learn=not args.no_auto_learn,
        )
        remembered_count = write_excel(
            transactions,
            output_path,
            category_memory=category_memory,
            account_name=args.account_name,
            currency_code=args.currency_code,
            currency_rate=args.currency_rate,
            overwrite=args.overwrite,
        )
        verify_excel(
            output_path,
            len(transactions),
            category_memory,
            args.account_name,
            args.currency_code,
            args.currency_rate,
        )

        income_count = sum(item.transaction_type == "Income" for item in transactions)
        expense_count = len(transactions) - income_count
        print(f"Готово: {output_path}")
        print(
            f"Операций: {len(transactions)} "
            f"(Income: {income_count}, Expense: {expense_count})."
        )
        print(
            f"Память категорий: {len(category_memory)} правил, "
            f"новых/обновленных: {learned_count}, применено: {remembered_count}."
        )
        return 0
    except (FileNotFoundError, FileExistsError, StatementParseError, ValueError) as exc:
        print(f"Ошибка: {exc}", file=sys.stderr)
        return 1


if __name__ == "__main__":
    raise SystemExit(main())
