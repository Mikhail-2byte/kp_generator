#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
Быстрая проверка работоспособности API ключа OpenRouter.

Использование:
    python scripts/check_api_key.py
    python scripts/check_api_key.py --key sk-or-v1-...
"""

import argparse
import os
import sys
from pathlib import Path

# Настраиваем кодировку UTF-8 для Windows
if sys.platform == 'win32':
    try:
        import codecs
        if sys.stdout.encoding != 'utf-8':
            sys.stdout = codecs.getwriter('utf-8')(sys.stdout.buffer, 'strict')
            sys.stderr = codecs.getwriter('utf-8')(sys.stderr.buffer, 'strict')
    except Exception:
        pass

# Добавляем корень проекта в путь
BASE_DIR = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(BASE_DIR))

import requests
from dotenv import load_dotenv

# Загружаем переменные окружения
load_dotenv()


def check_api_key(api_key: str) -> bool:
    """
    Проверяет работоспособность API ключа OpenRouter.
    
    Args:
        api_key: API ключ для проверки
        
    Returns:
        True если ключ работает, False иначе
    """
    url = "https://openrouter.ai/api/v1/chat/completions"
    
    headers = {
        "Authorization": f"Bearer {api_key}",
        "Content-Type": "application/json",
    }
    
    # Минимальный тестовый запрос
    payload = {
        "model": "xiaomi/mimo-v2-flash:free",  # Бесплатная модель для теста
        "messages": [
            {"role": "user", "content": "Привет"}
        ],
        "max_tokens": 10
    }
    
    try:
        print("Отправка тестового запроса к OpenRouter API...")
        response = requests.post(url, headers=headers, json=payload, timeout=10)
        
        if response.status_code == 200:
            print("[OK] API ключ работает!")
            data = response.json()
            if 'choices' in data and len(data['choices']) > 0:
                content = data['choices'][0].get('message', {}).get('content', '')
                print(f"[OK] Получен ответ от API: {content[:50]}...")
            return True
        elif response.status_code == 401:
            print("[ERROR] Ошибка 401: API ключ недействителен или отозван")
            try:
                error_data = response.json()
                if 'error' in error_data:
                    print(f"  Детали: {error_data['error'].get('message', 'Неизвестная ошибка')}")
            except Exception:
                pass
            return False
        elif response.status_code == 403:
            print("[ERROR] Ошибка 403: Недостаточно прав или баланса на аккаунте")
            print("  Пополните баланс на https://openrouter.ai/")
            return False
        elif response.status_code == 429:
            print("[ERROR] Ошибка 429: Превышен лимит запросов (rate limit)")
            retry_after = response.headers.get('Retry-After', '60')
            print(f"  Попробуйте через {retry_after} секунд")
            return False
        else:
            print(f"[ERROR] Ошибка {response.status_code}: {response.text[:200]}")
            return False
            
    except requests.exceptions.Timeout:
        print("[ERROR] Превышено время ожидания ответа от API")
        return False
    except requests.exceptions.ConnectionError:
        print("[ERROR] Не удалось подключиться к OpenRouter API")
        print("  Проверьте подключение к интернету")
        return False
    except Exception as e:
        print(f"[ERROR] Неожиданная ошибка: {e}")
        return False


def main():
    """Главная функция."""
    parser = argparse.ArgumentParser(
        description="Проверка работоспособности API ключа OpenRouter"
    )
    parser.add_argument(
        '--key',
        type=str,
        help='API ключ для проверки (если не указан, берется из OPENROUTER_API_KEY)'
    )
    
    args = parser.parse_args()
    
    # Получаем ключ
    api_key = args.key or os.getenv("OPENROUTER_API_KEY")
    
    if not api_key:
        print("[ERROR] Ошибка: API ключ не найден!")
        print("\nУкажите ключ одним из способов:")
        print("  1. Передайте через --key: python scripts/check_api_key.py --key sk-or-v1-...")
        print("  2. Установите переменную окружения: set OPENROUTER_API_KEY=sk-or-v1-...")
        print("  3. Добавьте в .env файл: OPENROUTER_API_KEY=sk-or-v1-...")
        sys.exit(1)
    
    # Проверяем формат
    if not api_key.startswith("sk-or-v1-"):
        print("[ERROR] Неверный формат API ключа!")
        print("  Ключ должен начинаться с 'sk-or-v1-'")
        print(f"  Получено: {api_key[:20]}...")
        sys.exit(1)
    
    print(f"Проверка API ключа: {api_key[:20]}...")
    print("-" * 50)
    
    # Проверяем работоспособность
    is_valid = check_api_key(api_key)
    
    print("-" * 50)
    if is_valid:
        print("[OK] Ключ работает корректно!")
        sys.exit(0)
    else:
        print("[ERROR] Ключ не работает или недоступен")
        sys.exit(1)


if __name__ == "__main__":
    main()

