# Промпт для реализации множественных позиций в Word шаблоне

## Краткое резюме

**Задача:** Добавить поддержку множественных позиций в Word документы, аналогично уже реализованной функциональности для Excel.

**Текущее состояние:**
- ✅ Excel: Класс `MultiPositionProcessor` успешно обрабатывает множественные позиции
- ❌ Word: Функция `generate_word_document()` работает только с одной позицией

**Что нужно:**
1. Создать класс `WordMultiPositionProcessor` для обработки множественных позиций в Word
2. Обновить `generate_word_document()` для использования нового класса
3. Обновить вызов в `routes/main.py` для передачи `positions` и `position_prices`

**Технологии:** Python, Flask, python-docx, паттерны SOLID и KISS

---

## Контекст проекта

Я работаю над Flask-приложением для генерации коммерческих предложений (КП). Система уже реализовала поддержку множественных позиций для Excel файлов, и теперь нужно добавить аналогичную функциональность для Word документов.

## Текущая ситуация

### Что уже работает (Excel):
- Класс `MultiPositionProcessor` в `app/services/multi_position_processor.py` успешно обрабатывает множественные позиции
- Автоматически добавляет новые строки в таблицу Excel
- Копирует стили, формулы и форматирование
- Обновляет итоговые расчеты и формулы
- Работает с любым количеством позиций

### Что нужно реализовать (Word):
- Функция `generate_word_document()` в `app/document_generator.py` сейчас работает только с одной позицией
- Нужно добавить поддержку множественных позиций аналогично Excel
- Нужно динамически добавлять строки в таблицу Word шаблона
- Сохранять форматирование и стили при добавлении новых строк

## Структура кода

### Текущая реализация Word (одна позиция):

Полный код находится в `app/document_generator.py`, строки 34-92:

```python
def generate_word_document(
    template_path,
    form_data,
    final_price,
    general_prise,
    final_price_NDS,
):
    """Формирует коммерческое предложение в формате Word на основе шаблона."""
    doc = Document(template_path)

    current_date = datetime.now().strftime('%d.%m.%Yг.')
    company = form_data['company'].strip()
    product = form_data['product'].strip()  # Только первая позиция!
    quantity = int(form_data['quantity'])
    cost_price = float(form_data['cost_price'])
    weight = float(form_data['weight'])
    logistics = float(form_data['logistics'])
    delivery_time = int(form_data['delivery_time'])
    tender_number = form_data.get('tender_number', '').strip()
    drawing_number = form_data.get('drawing_number', '').strip()
    material = form_data.get('material', '').strip()
    delivery_address = form_data.get('delivery_address', '').strip()
    duty_percent = float(form_data.get('duty_percent', 0))

    word_data = {
        '{{ company }}': company,
        '{{ product }}': product,
        '{{ quantity }}': str(quantity),
        '{{ cost_price }}': f"{cost_price:.0f}",
        '{{ weight }}': f"{weight:.0f}",
        '{{ logistics }}': f"{logistics:.0f}",
        '{{ final_price }}': f"{final_price:.0f}",
        '{{ general_prise }}': f"{general_prise:.0f}",
        '{{ final_price_NDS }}': f"{final_price_NDS:.0f}",
        '{{ tender_number }}': tender_number,
        '{{ drawing_number }}': drawing_number,
        '{{ material }}': material,
        '{{ delivery_address }}': delivery_address,
        '{{ date }}': current_date,
        '{{ duty_percent }}': f"{duty_percent:.1f}",
        '{{ delivery_time }}': str(delivery_time),
    }

    # Заменяет плейсхолдеры в параграфах
    for paragraph in doc.paragraphs:
        for key, value in word_data.items():
            if key in paragraph.text:
                paragraph.text = paragraph.text.replace(key, value)

    # Заменяет плейсхолдеры в таблицах
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for key, value in word_data.items():
                    if key in cell.text:
                        cell.text = cell.text.replace(key, value)

    word_file = BytesIO()
    doc.save(word_file)
    word_file.seek(0)
    return word_file
```

**Проблема:** Функция использует только первую позицию из `form_data` (например, `form_data['product']`), игнорируя множественные позиции.

### Реализация Excel (множественные позиции):

```python
def generate_excel_document(
    template_path,
    form_data,
    final_price,
    general_prise,
    position_prices=None,
):
    """Готовит Excel-файл в памяти."""
    positions = extract_positions_from_form(form_data)
    
    processor = MultiPositionProcessor(template_path)
    return processor.process_multiple_positions(
        positions,
        form_data,
        final_price,
        general_prise,
        position_prices=position_prices,
    )
```

### Структура данных позиций:

```python
position = {
    'product': 'Наименование товара',
    'drawing_number': 'Номер чертежа',
    'material': 'Материал',
    'cost_price': '1000',  # Сумма закупа
    'cost_price_per_kg': '100',  # Цена за кг
    'quantity': '5',  # Количество
    'weight': '10',  # Вес за шт (кг)
    'duty_percent': '5'  # Пошлина (%)
}

# position_prices содержит рассчитанные цены (список словарей):
position_price = {
    'position': {  # Исходные данные позиции
        'product': 'Наименование товара',
        'drawing_number': 'Номер чертежа',
        'material': 'Материал',
        'cost_price': '1000',
        'cost_price_per_kg': '100',
        'quantity': '5',
        'weight': '10',
        'duty_percent': '5'
    },
    'final_price': 1500.0,  # Цена за единицу (float)
    'general_price': 7500.0,  # Общая цена позиции (float) = final_price * quantity
    'costs': {  # Детализация затрат
        'cost_per_unit': 1200.0,
        'total_cost': 6000.0,
        # ... другие поля затрат
    },
    'margin': 20.0  # Маржа в процентах (float)
}
```

### Как вызывается в routes/main.py (строки 326-398):

```python
# Извлечение позиций
positions = validation.positions or extract_positions_from_form(form_data)

# Расчет цен для позиций
if len(positions) == 1:
    result = calculator.calculate_legacy_single_position(...)
    position_prices = [result]
    total_general_price = result['general_price']
else:
    calculation_result = calculator.calculate_multi_position_prices(...)
    position_prices = calculation_result['positions']
    total_general_price = calculation_result['total_revenue']

# Для совместимости используем первую позицию
first_position = position_prices[0]
final_price = first_position['final_price']
general_price = first_position['general_price']
final_price_nds = total_general_price * 1.2

# Генерация документов
excel_file = generate_excel_document(
    excel_template_path,
    form_data,
    final_price,
    total_general_price,
    position_prices=position_prices,  # ✅ Передается список позиций
)
word_file = generate_word_document(
    word_template_path, 
    form_data,  # ❌ Передается только form_data, без positions и position_prices
    final_price, 
    total_general_price, 
    final_price_nds
)
```

**Проблема:** В `generate_word_document()` не передаются `positions` и `position_prices`, поэтому функция не может обработать множественные позиции.

## Структура Word шаблона

Word шаблон (`templates_docs/template.docx`) содержит:
- Плейсхолдеры в формате `{{ field_name }}` в параграфах и ячейках таблиц
- Таблицу с позициями (предположительно одна строка-шаблон для позиции)
- Общие поля: `{{ company }}`, `{{ date }}`, `{{ logistics }}`, `{{ delivery_time }}`, и т.д.
- Поля позиции: `{{ product }}`, `{{ quantity }}`, `{{ cost_price }}`, и т.д.

### Полный список плейсхолдеров (из кода):

**Общие поля:**
- `{{ company }}` - Название компании
- `{{ date }}` - Текущая дата (формат: `dd.mm.YYYYг.`)
- `{{ logistics }}` - Стоимость логистики (целое число)
- `{{ delivery_time }}` - Срок поставки в днях
- `{{ tender_number }}` - Номер тендера
- `{{ delivery_address }}` - Адрес доставки
- `{{ final_price }}` - Цена за единицу (целое число)
- `{{ general_prise }}` - Общая цена (целое число)
- `{{ final_price_NDS }}` - Цена с НДС (целое число)

**Поля позиции:**
- `{{ product }}` - Наименование товара
- `{{ quantity }}` - Количество (целое число)
- `{{ cost_price }}` - Сумма закупа (целое число)
- `{{ weight }}` - Вес за шт (целое число)
- `{{ drawing_number }}` - Номер чертежа
- `{{ material }}` - Материал
- `{{ duty_percent }}` - Пошлина в процентах (формат: `X.X`)

**Примечание:** В текущей реализации все числовые значения форматируются как целые числа (`.0f`), кроме `duty_percent` (`.1f`).

## Требования к реализации

1. **Создать класс `WordMultiPositionProcessor`** аналогично `MultiPositionProcessor` для Excel
   - Должен находиться в `app/services/word_multi_position_processor.py`
   - Использовать библиотеку `python-docx` (уже установлена)

2. **Основной функционал:**
   - Определять таблицу с позициями в Word документе
   - Находить строку-шаблон для позиции (возможно, вторая строка таблицы, если первая - заголовок)
   - Добавлять новые строки в таблицу для каждой дополнительной позиции
   - Копировать стили и форматирование из строки-шаблона
   - Заполнять данные каждой позиции в соответствующие ячейки
   - Обновлять общие поля (компания, дата, итоговые суммы и т.д.)

3. **Обработка плейсхолдеров:**
   - Заменять общие плейсхолдеры (`{{ company }}`, `{{ date }}`, и т.д.)
   - Заполнять данные позиций в таблице
   - Обрабатывать итоговые значения (общая цена, цена с НДС)

4. **Обратная совместимость:**
   - Должна работать с одной позицией (как сейчас)
   - Автоматически определять количество позиций
   - Плавный переход между одной и множественными позициями

5. **Обновить `generate_word_document()`:**
   - Принимать параметры `positions` и `position_prices` (опционально)
   - Использовать `WordMultiPositionProcessor` для множественных позиций
   - Сохранять совместимость со старым кодом

6. **Обновить вызов в `routes/main.py`:**
   - Передавать `positions` и `position_prices` в `generate_word_document()`

## Технические детали

- **Библиотека**: `python-docx` (уже используется в проекте)
- **Стиль кода**: Следовать SOLID и KISS принципам
- **Обработка ошибок**: Корректная обработка отсутствующих таблиц, пустых данных и т.д.
- **Форматирование чисел**: Округление до целых для цен, форматирование процентов
- **Структура таблицы**: Нужно определить, какая строка является шаблоном (возможно, вторая строка, если первая - заголовок)

## Ожидаемый результат

1. **Детальный план реализации** с пошаговыми инструкциями
2. **Полный код класса `WordMultiPositionProcessor`** с комментариями
3. **Обновленный код `generate_word_document()`**
4. **Обновленный код вызова в `routes/main.py`**
5. **Рекомендации по структуре Word шаблона** (если нужны изменения)

## Дополнительные требования

- Код должен быть чистым, читаемым и поддерживаемым
- Следовать паттернам, используемым в `MultiPositionProcessor`
- Добавить docstrings на русском языке
- Обработать edge cases (пустые позиции, отсутствие таблиц, и т.д.)
- Сохранить все существующие функции (замена плейсхолдеров в параграфах)

## Вопросы для уточнения (если нужны):

1. Какая строка в таблице Word является шаблоном для позиции? (первая, вторая, последняя?)
2. Есть ли в таблице заголовок? Если да, то в какой строке?
3. Какие столбцы в таблице соответствуют каким полям позиции?
4. Нужно ли обновлять итоговую строку в таблице Word (если она есть)?

---

**Пожалуйста, предоставь:**
1. Детальный план реализации с учетом всех требований
2. Полный готовый код для интеграции
3. Примеры использования и тестирования

