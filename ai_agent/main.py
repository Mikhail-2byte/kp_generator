"""
Точка входа для запуска AI агента через терминал.
"""

import sys
from typing import Optional

from ai_agent.agent import AIAgent
from ai_agent.config import get_api_key


def print_help() -> None:
    """Выводит справку по командам."""
    print("""
Доступные команды:
  help, ?          - Показать эту справку
  clear, cls       - Очистить историю диалога
  exit, quit, выход - Выйти из программы

Примеры вопросов:
  - "Как обработать тендер?"
  - "Как работает расчет логистики?"
  - "Рассчитай логистику для 1000 кг в Москву"
  - "Сколько стоит доставка 2 тонн в Екатеринбург?"
""")


def setup_encoding() -> None:
    """Настраивает кодировку UTF-8 для Windows."""
    if sys.platform == 'win32':
        try:
            import codecs
            sys.stdout = codecs.getwriter('utf-8')(sys.stdout.buffer, 'strict')
            sys.stderr = codecs.getwriter('utf-8')(sys.stderr.buffer, 'strict')
        except Exception:
            pass  # Если не получилось, продолжаем


def main() -> None:
    """Основная функция запуска агента."""
    setup_encoding()
    
    print("=" * 60)
    print("AI Консультант для проекта KP Generator")
    print("=" * 60)
    print("\nИнициализация агента...")
    
    # Проверяем API ключ
    api_key = get_api_key()
    if not api_key:
        print("ОШИБКА: Не указан OPENROUTER_API_KEY!")
        print("Установите переменную окружения OPENROUTER_API_KEY или добавьте её в .env файл")
        sys.exit(1)
    
    try:
        agent = AIAgent()
        print("✓ Агент инициализирован")
        print("✓ База знаний загружена")
        print("\nГотов к работе! Введите 'help' для справки или 'exit' для выхода.\n")
    except Exception as e:
        print(f"ОШИБКА при инициализации: {e}")
        sys.exit(1)
    
    # Основной цикл
    while True:
        try:
            user_input = input("Вы: ").strip()
            
            if not user_input:
                continue
            
            # Обработка команд
            command = user_input.lower()
            
            if command in ['exit', 'quit', 'выход']:
                print("\nДо свидания!")
                break
            
            if command in ['help', '?']:
                print_help()
                continue
            
            if command in ['clear', 'cls']:
                agent.clear_history()
                print("История диалога очищена.\n")
                continue
            
            # Отправляем запрос агенту
            print("\nАгент думает...")
            response = agent.chat(user_input)
            print(f"\nАгент: {response}\n")
            
        except KeyboardInterrupt:
            print("\n\nПрервано пользователем. Выход...")
            break
        except EOFError:
            print("\n\nВыход...")
            break
        except Exception as e:
            print(f"\nОШИБКА: {e}\n")


if __name__ == "__main__":
    main()

