def calculate_selling_price(quantity, purchase_cost, duty_rate, logistics_cost_rub, weight, deal_length_days=170):
    # Константы из бюджета
    CONVERSION_RATE = 12  # Курс юаня к рублю
    LOGISTICS_CNR_RATIO = 0.3  # Доля логистики КНР
    LOGISTICS_RF_RATIO = 0.7  # Доля логистики РФ
    CONVERSION_FEE_RATE = 0.032  # Комиссия за конвертацию 3.2%
    CREDIT_RATE = 0.16  # Ставка кредита 16%
    VAT_RATE = 0.2  # НДС 20%
    
    # Перевод логистики в юани
    logistics_total_yuan = logistics_cost_rub / CONVERSION_RATE
    
    # Расчет логистики на единицу товара
    logistics_cnr_per_unit = (logistics_total_yuan * LOGISTICS_CNR_RATIO) / quantity
    logistics_rf_per_unit = (logistics_total_yuan * LOGISTICS_RF_RATIO) / quantity
    
    # Расчет пошлины на единицу товара
    duty_per_unit = (purchase_cost + logistics_cnr_per_unit) * duty_rate
    
    # Расчет стоимости конвертации
    conversion_fee = purchase_cost * quantity * CONVERSION_FEE_RATE
    
    # Расчет кредитных затрат
    credit_cost = purchase_cost * quantity * CREDIT_RATE / 365 * deal_length_days
    
    # Общие затраты на весь товар
    total_cost = (
        purchase_cost * quantity +  # Стоимость товара
        conversion_fee +  # Конвертация
        duty_per_unit * quantity +  # Пошлина
        logistics_cnr_per_unit * quantity +  # Логистика КНР
        logistics_rf_per_unit * quantity +  # Логистика РФ
        credit_cost  # Кредитные затраты
    )
    
    # Расчет цены для маржи 30% от выручки без НДС
    # Учитываем, что выручка без НДС = цена * количество
    # Маржа = (Выручка - Затраты) / Выручка = 0.3
    # => Выручка = Затраты / (1 - 0.3)
    revenue_without_vat = total_cost / (1 - 0.3)
    selling_price_per_unit = revenue_without_vat / quantity
    
    return selling_price_per_unit

# Пример использования
quantity = 1
purchase_cost = 35000  # Закупочная стоимость в юанях
duty_rate = 0.10  # Пошлина 5%
logistics_cost_rub = 150000  # Общая логистика в рублях
weight = 200  # Вес единицы товара
deal_length_days = 170  # Длина сделки

price = calculate_selling_price(quantity, purchase_cost, duty_rate, logistics_cost_rub, weight, deal_length_days)
print(f"Рассчитанная цена за единицу: {price:.2f} юаней")