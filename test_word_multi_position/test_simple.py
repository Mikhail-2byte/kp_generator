"""
Упрощенный тест для проверки работы WordMultiPositionProcessor.
Запустите в активированном виртуальном окружении.
"""
import sys
import os

# Добавляем корневую директорию в путь
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

# Прямой импорт без загрузки всего приложения
from docx import Document
from io import BytesIO
from datetime import datetime

# Импортируем только нужный класс
import importlib.util
spec = importlib.util.spec_from_file_location(
    "word_multi_position_processor",
    os.path.join(os.path.dirname(os.path.dirname(__file__)), 
                 "app", "services", "word_multi_position_processor.py")
)
module = importlib.util.module_from_spec(spec)
spec.loader.exec_module(module)
WordMultiPositionProcessor = module.WordMultiPositionProcessor


def test_basic():
    """Базовый тест."""
    print("=" * 60)
    print("ТЕСТ: Базовая проверка WordMultiPositionProcessor")
    print("=" * 60)
    
    template_path = os.path.join(os.path.dirname(os.path.dirname(__file__)), 
                                 'templates_docs', 'template.docx')
    
    if not os.path.exists(template_path):
        print(f"❌ Шаблон не найден: {template_path}")
        return False
    
    print(f"✅ Шаблон найден: {template_path}")
    
    # Тестовые данные
    form_data = {
        'company': 'Тестовая компания',
        'product': 'Тестовый товар',
        'quantity': '5',
        'cost_price': '1000',
        'weight': '10',
        'logistics': '5000',
        'delivery_time': '30',
        'tender_number': 'ТЕНДЕР-001',
        'drawing_number': 'ЧЕРТЕЖ-001',
        'material': 'Сталь',
        'delivery_address': 'Москва, ул. Тестовая, 1',
        'duty_percent': '5',
    }
    
    positions = [{
        'product': form_data['product'],
        'quantity': form_data['quantity'],
        'cost_price': form_data['cost_price'],
        'weight': form_data['weight'],
        'drawing_number': form_data['drawing_number'],
        'material': form_data['material'],
        'duty_percent': form_data['duty_percent'],
    }]
    
    position_prices = [{
        'final_price': 1500.0,
        'general_price': 7500.0,
    }]
    
    try:
        processor = WordMultiPositionProcessor(template_path)
        
        # Проверяем поиск таблицы
        table = processor.find_positions_table()
        if table is None:
            print("❌ Не найдена таблица с позициями")
            return False
        print(f"✅ Таблица с позициями найдена: {len(table.rows)} строк, {len(table.columns)} столбцов")
        
        # Проверяем поиск строки-шаблона
        template_row_idx = processor.find_template_row(table)
        if template_row_idx is None:
            print("❌ Не найдена строка-шаблон")
            return False
        print(f"✅ Строка-шаблон найдена: индекс {template_row_idx}")
        
        # Обрабатываем
        result = processor.process_multiple_positions(
            positions,
            form_data,
            final_price=1500.0,
            general_price=7500.0,
            final_price_nds=9000.0,
            position_prices=position_prices
        )
        
        if result is None:
            print("❌ Результат обработки пустой")
            return False
        
        # Сохраняем результат
        output_path = os.path.join(os.path.dirname(__file__), 'test_output_simple.docx')
        with open(output_path, 'wb') as f:
            f.write(result.getvalue())
        
        print(f"✅ Документ успешно создан: {output_path}")
        print(f"   Размер файла: {os.path.getsize(output_path)} байт")
        
        return True
        
    except Exception as e:
        print(f"❌ Ошибка: {e}")
        import traceback
        traceback.print_exc()
        return False


if __name__ == "__main__":
    success = test_basic()
    if success:
        print("\n✅ Тест пройден успешно!")
        sys.exit(0)
    else:
        print("\n❌ Тест провален")
        sys.exit(1)

