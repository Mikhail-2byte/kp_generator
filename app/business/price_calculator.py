from typing import Dict, Any, Optional

from app.core.exceptions import ValidationError


def calculate_selling_price(
    quantity: int,
    purchase_cost: float,
    logistics_rub: float,
    duty_percent: float,
    weight: float,
    delivery_time: int,
    margin_percent: float = 30,
    config: Optional[Dict[str, Any]] = None
) -> float:
    """
    Выполняет расчет продажной цены с учетом всех параметров бюджета.
    
    Рассчитывает финальную цену за единицу товара с учетом:
    - стоимости закупа
    - логистики (КНР и РФ)
    - пошлины
    - комиссии за конвертацию валюты
    - кредитных затрат
    - целевой маржинальности
    
    Args:
        quantity: Количество единиц товара
        purchase_cost: Стоимость закупа за единицу товара (в юанях)
        logistics_rub: Общая стоимость логистики (в рублях)
        duty_percent: Процент пошлины
        weight: Вес одной единицы товара (в кг)
        delivery_time: Время доставки (в днях)
        margin_percent: Целевая маржа в процентах (по умолчанию 30%)
        config: Словарь с конфигурацией расчета (курсы, коэффициенты и т.д.)
    
    Returns:
        Продажная цена за единицу товара (в рублях)
    
    Raises:
        ValidationError: Если quantity <= 0 или weight <= 0
    """
    # Проверка на нулевые значения
    if quantity <= 0:
        raise ValidationError(
            "Количество должно быть положительным числом",
            field='quantity',
            value=quantity
        )
    if weight <= 0:
        raise ValidationError(
            "Вес должен быть положительным числом",
            field='weight',
            value=weight
        )
    
    # Получаем константы из конфига или используем значения по умолчанию
    calc_config = config.get('calculation_constants', {}) if config else {}
    
    CONVERSION_RATE = calc_config.get('conversion_rate', 12)  # Курс юаня к рублю
    LOGISTICS_CNR_RATIO = calc_config.get('logistics_cnr_ratio', 0.3)  # Доля логистики КНР
    LOGISTICS_RF_RATIO = calc_config.get('logistics_rf_ratio', 0.7)  # Доля логистики РФ
    CONVERSION_FEE_RATE = calc_config.get('conversion_fee_rate', 0.032)  # Комиссия за конвертацию 3.2%
    CREDIT_RATE = calc_config.get('credit_rate', 0.16)  # Ставка кредита 16%
    
    # Расчет общего веса
    total_weight = weight * quantity
    
    # Перевод логистики в юани и распределение по весу
    logistics_total_yuan = logistics_rub / CONVERSION_RATE
    
    # Расчет логистики на единицу товара (пропорционально весу)
    logistics_cnr_per_unit = (logistics_total_yuan * LOGISTICS_CNR_RATIO * weight) / total_weight
    logistics_rf_per_unit = (logistics_total_yuan * LOGISTICS_RF_RATIO * weight) / total_weight
    
    # Расчет пошлины на единицу товара
    duty_per_unit = (purchase_cost + logistics_cnr_per_unit) * (duty_percent / 100)
    
    # Расчет стоимости конвертации
    conversion_fee = purchase_cost * quantity * CONVERSION_FEE_RATE
    conversion_fee_per_unit = conversion_fee / quantity
    
    # Расчет кредитных затрат
    credit_cost = purchase_cost * quantity * CREDIT_RATE / 365 * delivery_time
    credit_cost_per_unit = credit_cost / quantity
    
    # Общие затраты на единицу товара
    total_cost_per_unit = (
        purchase_cost +
        logistics_cnr_per_unit +
        logistics_rf_per_unit +
        duty_per_unit +
        conversion_fee_per_unit +
        credit_cost_per_unit
    )
    
    # Расчет цены для маржи margin_percent%
    selling_price_per_unit = total_cost_per_unit / (1 - margin_percent / 100)
    
    return selling_price_per_unit
