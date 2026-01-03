"""
Скрипт запуска AI агента.
Используется для обхода проблем с кодировкой при запуске через -m.
"""

import os
import sys

# Настраиваем кодировку UTF-8 для Windows
if sys.platform == 'win32':
    try:
        import codecs
        sys.stdout = codecs.getwriter('utf-8')(sys.stdout.buffer, 'strict')
        sys.stderr = codecs.getwriter('utf-8')(sys.stderr.buffer, 'strict')
    except Exception:
        pass

# Добавляем корень проекта в путь
BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
if BASE_DIR not in sys.path:
    sys.path.insert(0, BASE_DIR)

# Запускаем main
if __name__ == "__main__":
    from ai_agent.main import main
    main()

