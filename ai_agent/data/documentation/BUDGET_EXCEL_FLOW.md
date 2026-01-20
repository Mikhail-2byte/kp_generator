# Как заполняется бюджет Excel и откуда берется продажная цена

## Как запустить тесты

### Вариант 1: Через pytest напрямую
```bash
# Активируйте виртуальное окружение (если используется)
# Windows:
venv\Scripts\activate

# Затем запустите тесты:
pytest tests/generation/test_budget_creation.py -v

# Или все тесты:
pytest tests/ -v

# С покрытием:
pytest tests/ --cov=app --cov-report=html
```

### Вариант 2: Через Python модуль
```bash
python -m pytest tests/generation/test_budget_creation.py -v
```

### Вариант 3: Конкретный тест
```bash
pytest tests/generation/test_budget_creation.py::TestBudgetCreation::test_budget_creation_single_position -v
```

---

## Полный путь продажной цены: от расчета до Excel

### 1. Начало: Расчет продажной цены

**Файл:** `app/business/price_calculator.py`

**Функция:** `calculate_selling_price()`

**Формула расчета:**
```python
# Общие затраты на единицу товара
total_cost_per_unit = (
    purchase_cost +                    # Стоимость закупа (в юанях)
    logistics_cnr_per_unit +          # Логистика КНР на единицу
    logistics_rf_per_unit +           # Логистика РФ на единицу
    duty_per_unit +                   # Пошлина на единицу
    conversion_fee_per_unit +         # Комиссия за конвертацию валюты
    credit_cost_per_unit              # Кредитные затраты
)

# Продажная цена с учетом маржи
selling_price_per_unit = total_cost_per_unit / (1 - margin_percent / 100)
```

**Входные параметры:**
- `quantity` - количество единиц товара
- `purchase_cost` - стоимость закупа за единицу (в юанях)
- `logistics_rub` - общая стоимость логистики (в рублях)
- `duty_percent` - процент пошлины
- `weight` - вес одной единицы товара (в кг)
- `delivery_time` - время доставки (в днях)
- `margin_percent` - целевая маржа в процентах (по умолчанию 30%)

**Выход:** Продажная цена за единицу товара (в рублях)

---

### 2. Оркестрация расчета

**Файл:** `app/services/generation_orchestrator.py`

**Метод:** `calculate_prices()`

**Логика:**
- Для **одной позиции**: использует `calculate_legacy_single_position()` → вызывает `calculate_selling_price()`
- Для **множественных позиций**: использует `calculate_multi_position_prices()` → рассчитывает единую маржу для всех позиций

**Округление:**
```python
# Округляем final_price вверх до десятков
final_price_rounded = math.ceil(result['final_price'] / 10.0) * 10.0
```

**Результат:** 
- `position_prices` - список с рассчитанными ценами для каждой позиции
- `total_general_price` - общая цена всех позиций

---

### 3. Генерация Excel документа

**Файл:** `app/business/document_generator.py`

**Функция:** `generate_excel_document()`

**Вызов:**
```python
excel_file = generate_excel_document(
    template_path='templates_docs/template.xlsx',
    form_data=form_data,
    final_price=final_price,              # Цена первой позиции
    general_prise=total_general_price,     # Общая цена всех позиций
    position_prices=position_prices,       # Список цен по позициям
    manager_fio=manager_fio
)
```

---

### 4. Заполнение Excel файла

**Файл:** `app/services/multi_position_processor.py`

**Класс:** `MultiPositionProcessor`

#### 4.1. Общие данные (fill_common_data)

**Ячейки Excel:**
- `D2` - Дата (текущая дата)
- `D4` - Название компании (`form_data['company']`)
- `D5` - Номер тендера (`form_data['tender_number']`)
- `P4` - Адрес доставки (`form_data['delivery_address']`)
- `U14` - Логистика (`form_data['logistics']`)
- `I15` - Срок поставки (`form_data['delivery_time']`)
- `B16` - Условия оплаты (текст)
- `I16` - Условия оплаты (число дней)
- `N22` - ФИО менеджера (`manager_fio`)
- `H10` - **Цена за единицу** (округляется вверх до десятков)
- `I11` - **Общая цена** (округляется до целого числа)

**Код:**
```python
if final_price is not None:
    # Округляем до целого числа
    rounded_fp = round(float(final_price))
    sheet['H10'] = rounded_fp  # Цена за единицу

if general_price is not None:
    gp = round(float(general_price))  # Округление до целого
    sheet['I11'] = gp  # Общая цена
```

#### 4.2. Данные позиций (fill_position_data)

**Ячейки Excel (для каждой позиции, начиная со строки 10):**

| Столбец | Поле | Источник данных |
|---------|------|-----------------|
| B | Номер позиции | Автоматически: 1, 2, 3... |
| C | Наименование товара | `position['product']` + `position['drawing_number']` |
| D | Материал | `position['material']` |
| E | Номер чертежа | `position['drawing_number']` |
| G | Количество | `position['quantity']` |
| H | **Цена за единицу** | `position_prices[i]['final_price']` (округляется вверх до десятков) |
| I | Выручка | Формула: `=H{row}*G{row}` |
| M | Сумма закупа | `position['cost_price']` |
| N | Цена за кг | `position['cost_price_per_kg']` |
| P | Вес за шт | `position['weight']` |
| X | Пошлина | `position['duty_percent']` (в процентах) |

**Код заполнения цены:**
```python
# Проставляем рассчитанные цены по позиции
if position_prices and i < len(position_prices):
    pp = position_prices[i]
    fp = pp.get('final_price')
    if fp is not None:
        # Округляем до целого числа
        fp_rounded = round(float(fp))
        sheet[f"H{row_number}"] = fp_rounded  # Цена за единицу по позиции
    
    # Вставляем формулу для выручки: цена за единицу * количество
    sheet[f"I{row_number}"] = f"=H{row_number}*G{row_number}"
```

---

## Схема потока данных

```
1. Пользователь заполняет форму
   ↓
2. GenerationOrchestrator.orchestrate()
   ↓
3. Валидация данных (validate_request)
   ↓
4. Расчет цен (calculate_prices)
   ├─ Для одной позиции:
   │   └─ calculate_legacy_single_position()
   │       └─ calculate_selling_price() ← ОСНОВНОЙ РАСЧЕТ
   │
   └─ Для множественных позиций:
       └─ calculate_multi_position_prices()
           └─ calculate_position_costs() (для каждой позиции)
           └─ Расчет общего коэффициента для единой маржи
   ↓
5. Округление цен:
   - final_price округляется до целого числа: round(price)
   - general_price округляется до целого: round(price)
   ↓
6. Генерация Excel (generate_excel_document)
   ↓
7. MultiPositionProcessor.process_multiple_positions()
   ├─ fill_common_data() → Заполнение общих полей
   │   ├─ H10 = final_price (округлено до целого числа)
   │   └─ I11 = general_price (округлено до целого)
   │
   └─ fill_position_data() → Заполнение данных каждой позиции
       └─ H{row} = final_price позиции (округлено до целого числа)
       └─ I{row} = формула =H{row}*G{row} (выручка)
```

---

## Ключевые моменты

### Откуда берется продажная цена:

1. **Расчет:** `calculate_selling_price()` в `price_calculator.py`
   - Формула: `total_cost_per_unit / (1 - margin_percent / 100)`
   - Учитывает: закуп, логистику, пошлину, конвертацию, кредит, маржу

2. **Округление:** 
   - В `generation_orchestrator.py`: округляется до целого числа
   - Пример: 1234.56 → 1235

3. **Вставка в Excel:**
   - Для одной позиции: `H10` (общая цена за единицу)
   - Для каждой позиции: `H{row_number}` (цена за единицу позиции)
   - Выручка: `I{row_number}` = формула `=H{row}*G{row}`

### Где вставляются значения в Excel:

**Общие поля:**
- `D2` - Дата
- `D4` - Компания
- `D5` - Номер тендера
- `P4` - Адрес доставки
- `U14` - Логистика
- `I15` - Срок поставки
- `B16` - Условия оплаты (текст)
- `I16` - Условия оплаты (дни)
- `N22` - ФИО менеджера
- `H10` - Цена за единицу (для одной позиции)
- `I11` - Общая цена (для одной позиции)

**Поля позиций (начиная со строки 10, DATA_START_ROW = 10):**
- `B{row}` - Номер позиции (1, 2, 3...)
- `C{row}` - Наименование товара + номер чертежа
- `D{row}` - Материал
- `E{row}` - Номер чертежа
- `G{row}` - Количество
- `H{row}` - **Цена за единицу** (продажная цена, округленная вверх до десятков)
- `I{row}` - Выручка (формула =H{row}*G{row})
- `M{row}` - Сумма закупа
- `N{row}` - Цена за кг
- `P{row}` - Вес за шт
- `X{row}` - Пошлина (в процентах)

---

## Пример расчета

**Входные данные:**
- Количество: 10 шт
- Стоимость закупа: 1000 юаней/шт
- Логистика: 50000 рублей
- Пошлина: 5%
- Вес: 5 кг/шт
- Срок доставки: 30 дней
- Маржа: 30%

**Расчет:**
1. Общий вес: 10 × 5 = 50 кг
2. Логистика на единицу: (50000 / 12) × (5 / 50) = ~416.67 юаней
3. Пошлина: (1000 + 416.67) × 0.05 = ~70.83 юаней
4. Конвертация: 1000 × 0.032 = 32 юани
5. Кредит: 1000 × 0.16 / 365 × 30 = ~13.15 юаней
6. **Общие затраты:** 1000 + 416.67 + 416.67 + 70.83 + 32 + 13.15 = ~1949.32 юаней
7. **Продажная цена:** 1949.32 / (1 - 0.30) = ~2784.74 юаней
8. **Округление:** round(2784.74) = 2785 юаней

**В Excel:**
- `H10` (или `H{row}` для позиции) = 2785
- `I{row}` = формула `=H{row}*G{row}` = 2785 × 10 = 27850

