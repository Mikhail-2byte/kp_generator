# Установка и запуск

1. Создайте виртуальное окружение и активируйте его:
   ```powershell
   python -m venv venv
   .\venv\Scripts\Activate.ps1
   ```
2. Установите зависимости:
   ```powershell
   pip install -r requirements.txt
   ```
3. Скопируйте `.env.example` в `.env` и задайте необходимые значения.
   ```powershell
   copy .env.example .env
   ```
4. Примените миграции:
   ```powershell
   alembic upgrade head
   ```
5. Запустите приложение в режиме разработки:
   ```powershell
   python app.py
   ```
6. Для production-режима включите Waitress (в `.env` установите `USE_WAITRESS=1`) и запустите сервер:
   ```powershell
   waitress-serve --call "app:create_app"
   ```
7. Запустите тесты при необходимости:
   ```powershell
   pytest
   ```
