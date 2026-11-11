# PDF to Akaunting Excel

Flask-приложение читает банковские выписки OnPoint через `pdfplumber`, начиная со второй страницы, и создаёт Excel для импорта транзакций в Akaunting.

## Возможности

- загрузка PDF через защищённую страницу входа;
- чтение всех страниц выписки, начиная со второй;
- экспорт в чистый `.xlsx` без цветового оформления;
- последовательные номера `TRA-00001`, `TRA-00002` и далее;
- определение `income` и `expense`;
- запоминание вручную назначенных категорий в MySQL;
- повторное использование категорий по названию продавца после `Point Of Sale Withdrawal` или `Point Of Sale Deposit`;
- подсказки из уже созданных категорий.

## Развёртывание на Render

1. Загрузите содержимое этой папки в корень GitHub-репозитория.
2. В Render выберите **New Web Service → Git Provider** и подключите репозиторий.
3. Укажите:

```text
Language: Python
Build Command: pip install -r requirements.txt
Start Command: gunicorn --bind 0.0.0.0:$PORT --workers 1 --threads 2 --timeout 180 wsgi:app
Health Check Path: /health
```

4. Добавьте переменные окружения:

```text
LOGIN_USERNAME=startech
LOGIN_PASSWORD=замените-на-свой-пароль
SECRET_KEY=случайная-длинная-строка
MYSQL_HOST=адрес-mysql-сервера
MYSQL_PORT=3306
MYSQL_USER=пользователь-mysql
MYSQL_PASSWORD=пароль-mysql
MYSQL_DATABASE=pdf_converter
SESSION_COOKIE_SECURE=true
TRUST_PROXY=true
```

Render задаёт `PORT` автоматически. Приложение само создаст необходимые таблицы внутри указанной базы MySQL. Сервер MySQL должен разрешать подключения с хостинга Render.

> Render не предоставляет встроенную MySQL-базу. Нужна внешняя MySQL/MariaDB. На бесплатном Render локальные загруженные PDF и созданные Excel могут исчезнуть после перезапуска сервиса, но правила категорий сохранятся во внешней MySQL.

## Локальный запуск

Требуется Python 3.12 и доступный сервер MySQL.

```bash
python3 -m venv .venv
source .venv/bin/activate
python -m pip install -r requirements.txt
python web_app.py
```

После запуска откройте `http://127.0.0.1:5000`.

Локальные значения по умолчанию:

```text
Логин: startech
Пароль: 12345678
MySQL: root@127.0.0.1:3306/pdf_converter
```

Для хостинга обязательно задайте собственные `LOGIN_PASSWORD` и `SECRET_KEY` через переменные окружения.

## Формат Excel

Лист `Transactions` содержит колонки:

```text
type
number
paid_at
amount
currency_code
currency_rate
account_name
invoice_bill_number
contact_email
category_name
description
payment_method
```

Основные значения:

- `account_name`: `Business Account`;
- `currency_code`: `USD`;
- `currency_rate`: `1`;
- `payment_method`: `offline-payments.cash.1`;
- `category_name`: пустое для неизвестной категории;
- `invoice_bill_number` и `contact_email`: пустые.

## Структура проекта

- `web_app.py` — Flask-интерфейс, авторизация и работа с MySQL;
- `main.py` — чтение PDF и создание Excel;
- `wsgi.py` — точка запуска Gunicorn;
- `templates/` — HTML-шаблоны;
- `static/` — JavaScript и CSS;
- `requirements.txt` — зависимости Python;
- `Procfile` — команда запуска веб-сервиса.

Папка `instance/` создаётся автоматически и не должна попадать в Git.
