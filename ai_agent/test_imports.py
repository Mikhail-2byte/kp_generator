"""
Тестовый скрипт для проверки импортов и базовой функциональности.
"""

import sys
import os

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

print("Проверка импортов...")

try:
    print("1. Импорт requests...")
    import requests
    print(f"   ✓ requests {requests.__version__}")
except ImportError as e:
    print(f"   ✗ Ошибка: {e}")
    sys.exit(1)

try:
    print("2. Импорт config...")
    from ai_agent.config import get_api_key, get_model_name
    api_key = get_api_key()
    model = get_model_name()
    print(f"   ✓ API ключ: {'установлен' if api_key else 'НЕ установлен'}")
    print(f"   ✓ Модель: {model}")
except Exception as e:
    print(f"   ✗ Ошибка: {e}")
    sys.exit(1)

try:
    print("3. Импорт knowledge_base...")
    from ai_agent.knowledge_base import KnowledgeBase
    kb = KnowledgeBase()
    print(f"   ✓ Загружено инструкций: {len(kb.instructions)}")
    print(f"   ✓ Загружено документации: {len(kb.documentation)}")
except Exception as e:
    print(f"   ✗ Ошибка: {e}")
    import traceback
    traceback.print_exc()
    sys.exit(1)

try:
    print("4. Импорт logistics_helper...")
    from ai_agent.logistics_helper import LogisticsHelper
    lh = LogisticsHelper()
    cities = lh._load_cities()
    print(f"   ✓ Загружено городов: {len(cities)}")
except Exception as e:
    print(f"   ✗ Ошибка: {e}")
    import traceback
    traceback.print_exc()
    sys.exit(1)

try:
    print("5. Импорт agent...")
    from ai_agent.agent import AIAgent
    print("   ✓ AIAgent импортирован")
except Exception as e:
    print(f"   ✗ Ошибка: {e}")
    import traceback
    traceback.print_exc()
    sys.exit(1)

print("\n✓ Все импорты успешны!")
print("\nДля запуска агента используйте:")
print("  python ai_agent/run.py")

