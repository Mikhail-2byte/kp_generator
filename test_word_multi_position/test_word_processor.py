"""
Тестовый скрипт для проверки работы WordMultiPositionProcessor.
Запустите этот скрипт в окружении с установленным python-docx.
"""
import sys
import os
from io import BytesIO

# Добавляем корневую директорию в путь
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from app.services.word_multi_position_processor import WordMultiPositionProcessor
from app.helpers import extract_positions_from_form


def test_single_position():
    """Тест с одной позицией (проверка обратной совместимости)."""
    print("=" * 60)
    print("ТЕСТ 1: Одна позиция (обратная совместимость)")
    print("=" * 60)
    
    template_path = os.path.join(os.path.dirname(os.path.dirname(__file__)), 'templates_docs', 'template.docx')
    
    if not os.path.exists(template_path):
        print(f"❌ Шаблон не найден: {template_path}")
        return False
    
    # Подготовка тестовых данных (одна позиция)
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
    
    positions = extract_positions_from_form(form_data)
    
    position_prices = [{
        'position': positions[0],
        'final_price': 1500.0,
        'general_price': 7500.0,
        'costs': {},
        'margin': 20.0
    }]
    
    try:
        processor = WordMultiPositionProcessor(template_path)
        
        # Проверяем, что можем найти таблицу
        table = processor.find_positions_table()
        if table is None:
            print("❌ Не найдена таблица с позициями")
            return False
        print(f"✅ Таблица с позициями найдена: {len(table.rows)} строк, {len(table.columns)} столбцов")
        
        # Проверяем, что можем найти строку-шаблон
        template_row_idx = processor.find_template_row(table)
        if template_row_idx is None:
            print("❌ Не найдена строка-шаблон")
            return False
        print(f"✅ Строка-шаблон найдена: индекс {template_row_idx}")
        
        # Обрабатываем позиции
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
        
        # Сохраняем результат для проверки
        output_path = os.path.join(os.path.dirname(__file__), 'test_output_single.docx')
        with open(output_path, 'wb') as f:
            f.write(result.getvalue())
        
        print(f"✅ Документ успешно создан: {output_path}")
        print(f"   Размер файла: {os.path.getsize(output_path)} байт")
        return True
        
    except Exception as e:
        print(f"❌ Ошибка при обработке: {e}")
        import traceback
        traceback.print_exc()
        return False


def test_multiple_positions():
    """Тест с множественными позициями."""
    print("\n" + "=" * 60)
    print("ТЕСТ 2: Множественные позиции")
    print("=" * 60)
    
    template_path = os.path.join(os.path.dirname(os.path.dirname(__file__)), 'templates_docs', 'template.docx')
    
    if not os.path.exists(template_path):
        print(f"❌ Шаблон не найден: {template_path}")
        return False
    
    # Подготовка тестовых данных (несколько позиций)
    form_data = {
        'company': 'Тестовая компания',
        'product': 'Товар 1',
        'product_2': 'Товар 2',
        'product_3': 'Товар 3',
        'quantity': '5',
        'quantity_2': '10',
        'quantity_3': '15',
        'cost_price': '1000',
        'cost_price_2': '2000',
        'cost_price_3': '3000',
        'weight': '10',
        'weight_2': '20',
        'weight_3': '30',
        'logistics': '5000',
        'delivery_time': '30',
        'tender_number': 'ТЕНДЕР-002',
        'drawing_number': 'ЧЕРТЕЖ-001',
        'drawing_number_2': 'ЧЕРТЕЖ-002',
        'drawing_number_3': 'ЧЕРТЕЖ-003',
        'material': 'Сталь',
        'material_2': 'Алюминий',
        'material_3': 'Медь',
        'delivery_address': 'Москва, ул. Тестовая, 1',
        'duty_percent': '5',
        'duty_percent_2': '7',
        'duty_percent_3': '10',
    }
    
    positions = extract_positions_from_form(form_data)
    print(f"✅ Извлечено позиций: {len(positions)}")
    
    position_prices = [
        {
            'position': positions[0],
            'final_price': 1500.0,
            'general_price': 7500.0,
            'costs': {},
            'margin': 20.0
        },
        {
            'position': positions[1],
            'final_price': 2500.0,
            'general_price': 25000.0,
            'costs': {},
            'margin': 20.0
        },
        {
            'position': positions[2],
            'final_price': 3500.0,
            'general_price': 52500.0,
            'costs': {},
            'margin': 20.0
        }
    ]
    
    try:
        processor = WordMultiPositionProcessor(template_path)
        
        # Обрабатываем позиции
        result = processor.process_multiple_positions(
            positions,
            form_data,
            final_price=1500.0,
            general_price=85000.0,  # Общая сумма всех позиций
            final_price_nds=102000.0,
            position_prices=position_prices
        )
        
        if result is None:
            print("❌ Результат обработки пустой")
            return False
        
        # Сохраняем результат для проверки
        output_path = os.path.join(os.path.dirname(__file__), 'test_output_multiple.docx')
        with open(output_path, 'wb') as f:
            f.write(result.getvalue())
        
        print(f"✅ Документ успешно создан: {output_path}")
        print(f"   Размер файла: {os.path.getsize(output_path)} байт")
        
        # Проверяем, что в документе действительно несколько строк
        from docx import Document
        doc = Document(output_path)
        table = processor.find_positions_table()
        if table:
            print(f"   Строк в таблице позиций: {len(table.rows)}")
        
        return True
        
    except Exception as e:
        print(f"❌ Ошибка при обработке: {e}")
        import traceback
        traceback.print_exc()
        return False


if __name__ == "__main__":
    print("\n" + "=" * 60)
    print("ТЕСТИРОВАНИЕ WordMultiPositionProcessor")
    print("=" * 60 + "\n")
    
    # Проверяем наличие шаблона
    template_path = os.path.join(os.path.dirname(os.path.dirname(__file__)), 'templates_docs', 'template.docx')
    if not os.path.exists(template_path):
        print(f"❌ Шаблон не найден: {template_path}")
        print("Убедитесь, что файл template.docx существует в папке templates_docs/")
        sys.exit(1)
    
    # Запускаем тесты
    test1_result = test_single_position()
    test2_result = test_multiple_positions()
    
    print("\n" + "=" * 60)
    print("РЕЗУЛЬТАТЫ ТЕСТИРОВАНИЯ")
    print("=" * 60)
    print(f"Тест 1 (одна позиция): {'✅ ПРОЙДЕН' if test1_result else '❌ ПРОВАЛЕН'}")
    print(f"Тест 2 (множественные позиции): {'✅ ПРОЙДЕН' if test2_result else '❌ ПРОВАЛЕН'}")
    
    if test1_result and test2_result:
        print("\n✅ Все тесты пройдены успешно!")
        sys.exit(0)
    else:
        print("\n❌ Некоторые тесты провалены. Проверьте вывод выше.")
        sys.exit(1)

