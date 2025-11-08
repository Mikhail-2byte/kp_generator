"""
Базовый тест для проверки структуры Word документа.
Запустите этот скрипт в окружении с установленным python-docx.
"""
import sys
import os

# Добавляем корневую директорию в путь
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

def test_basic_import():
    """Проверяет, что можем импортировать необходимые модули."""
    try:
        from docx import Document
        print("✅ python-docx импортирован успешно")
        return True
    except ImportError as e:
        print(f"❌ Ошибка импорта: {e}")
        print("Установите python-docx: pip install python-docx")
        return False

def test_template_exists():
    """Проверяет наличие шаблона."""
    template_path = os.path.join(os.path.dirname(os.path.dirname(__file__)), 'templates_docs', 'template.docx')
    if os.path.exists(template_path):
        print(f"✅ Шаблон найден: {template_path}")
        return template_path
    else:
        print(f"❌ Шаблон не найден: {template_path}")
        return None

if __name__ == "__main__":
    print("Проверка окружения для работы с Word шаблонами\n")
    print("=" * 60)
    
    if test_basic_import():
        template_path = test_template_exists()
        if template_path:
            print("\n✅ Все проверки пройдены!")
            print("\nТеперь можно запустить analyze_template.py для анализа структуры шаблона.")
        else:
            print("\n⚠️  Шаблон не найден, но python-docx установлен.")
    else:
        print("\n❌ Необходимо установить зависимости.")

