# ПК МИШКА — Telegram Web App Генератор Договоров и Сертификатов

## Что внутри архива
- `main.py` — сервер FastAPI + Telegram-бот (Aiogram 3.x) с вебхуками и авто-пингером для Render.
- `document_builder.py` — модуль автозаполнения Дилерского договора (`.docx`) и Сертификата (`.pptx` + `.pdf` через Pillow).
- `static/index.html` — интерфейс Telegram Web App с анимациями и Glassmorphism.
- `templates/Дилерский_договор_Мишка-2_2.docx` — готовый отформатированный шаблон дилерского договора.
- `requirements.txt` и `Procfile` — конфигурация для деплоя на Render.

## Что нужно добавить перед пушем в GitHub
1. Положите ваш файл презентации сертификата `MISHKA_certificate_A4.pptx` в папку `templates/`.
2. *(Опционально для PDF-сертификата)*: положите чистый фон сертификата `cert_bg.png` и шрифт `font.ttf` в папку `templates/`.

## Переменные окружения на Render
- `BOT_TOKEN` — токен вашего бота от @BotFather
- `WEBHOOK_URL` — адрес сервиса на Render (например, `https://mishka-doc-bot.onrender.com`)
- `PYTHON_VERSION` — `3.11.9`
