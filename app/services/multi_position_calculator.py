# app/services/multi_position_calculator.py
from typing import Any, Dict, List, Optional

from app.business.interfaces import PriceCalculatorPort
from app.business.price_calculator import calculate_selling_price


class MultiPositionCalculator(PriceCalculatorPort):
    """
    Калькулятор для множественных позиций с единой итоговой маржой.
    
    Поддерживает два режима ценообразования:
    - 'global': общий коэффициент по целевой суммарной марже
    - 'per_position': каждая позиция достигает целевой маржи
    
    Attributes:
        config: Конфигурация приложения с параметрами расчета
        calc_config: Параметры расчета из конфигурации
        pricing_mode: Режим ценообразования ('global' или 'per_position')
    """
    
    def __init__(self, config: Optional[Dict[str, Any]] = None) -> None:
        """
        Инициализирует калькулятор.
        
        Args:
            config: Словарь конфигурации приложения
        """
        self.config = config or {}
        self.calc_config = self.config.get('calculation_constants', {})
        pricing_cfg = (self.config or {}).get('pricing', {})
        # Режим ценообразования: 'global' (по общему коэффициенту) или 'per_position' (каждая позиция к целевой марже)
        self.pricing_mode = pricing_cfg.get('mode', 'global')
        
        # Константы расчета
        self.CONVERSION_RATE = self.calc_config.get('conversion_rate', 11.5)
        self.LOGISTICS_CNR_RATIO = self.calc_config.get('logistics_cnr_ratio', 0.3)
        self.LOGISTICS_RF_RATIO = self.calc_config.get('logistics_rf_ratio', 0.7)
        self.CONVERSION_FEE_RATE = self.calc_config.get('conversion_fee_rate', 0.032)
        self.CREDIT_RATE = self.calc_config.get('credit_rate', 0.16)
        self.VAT_RATE = self.calc_config.get('vat_rate', 0.22)  # Ставка НДС 22%
    
    def calculate_position_costs(
        self, position: Dict[str, Any], logistics_rub: float, 
        delivery_time: int, total_weight: float,
        use_credit: bool = False,
        payment_days: Optional[int] = None
    ) -> Dict[str, float]:
        """
        Рассчитывает все затраты для одной позиции.
        
        Args:
            position: Словарь с данными позиции (quantity, cost_price, weight, duty_percent)
            logistics_rub: Общая стоимость логистики (в рублях)
            delivery_time: Время доставки (в днях)
            total_weight: Общий вес всех позиций (в кг)
            use_credit: Использовать ли кредит в расчете
            payment_days: Количество дней оплаты (для расчета кредита, как в Excel K15 = I15 + I16)
        
        Returns:
            Словарь с рассчитанными затратами:
            - cost_per_unit: Общие затраты на единицу товара
            - total_cost: Общие затраты на всю позицию
            - logistics_cnr_per_unit: Логистика КНР на единицу
            - logistics_rf_per_unit: Логистика РФ на единицу
            - duty_per_unit: Пошлина на единицу
            - conversion_fee_per_unit: Комиссия за конвертацию на единицу
            - credit_cost_per_unit: Кредитные затраты на единицу
        """
        quantity = int(position['quantity'])
        cost_price = float(position['cost_price'])
        weight = float(position['weight'])
        duty_percent = float(position.get('duty_percent', 0))
        
        # Общий вес позиции
        position_weight = weight * quantity
        
        # Пропорциональная логистика для позиции
        position_logistics = logistics_rub * (position_weight / total_weight) if total_weight > 0 else 0
        
        # Перевод логистики в юани
        logistics_total_yuan = position_logistics / self.CONVERSION_RATE
        
        # Расчет логистики на единицу товара
        if position_weight > 0:
            logistics_cnr_per_unit = (logistics_total_yuan * self.LOGISTICS_CNR_RATIO * weight) / position_weight
            logistics_rf_per_unit = (logistics_total_yuan * self.LOGISTICS_RF_RATIO * weight) / position_weight
        else:
            logistics_cnr_per_unit = 0
            logistics_rf_per_unit = 0
        
        # Расчет пошлины на единицу товара
        duty_per_unit = (cost_price + logistics_cnr_per_unit) * (duty_percent / 100)
        
        # Расчет стоимости конвертации
        conversion_fee = cost_price * quantity * self.CONVERSION_FEE_RATE
        conversion_fee_per_unit = conversion_fee / quantity if quantity > 0 else 0
        
        # Расчет кредитных затрат (только если use_credit=True)
        # В Excel формула: I28*16%/365*K15, где K15 = I15 + I16 (срок поставки + условия оплаты)
        if use_credit:
            # Используем delivery_time + payment_days, как в Excel (K15 = I15 + I16)
            credit_days = delivery_time + (payment_days if payment_days is not None else 0)
            credit_cost = cost_price * quantity * self.CREDIT_RATE / 365 * credit_days
            credit_cost_per_unit = credit_cost / quantity if quantity > 0 else 0
        else:
            credit_cost_per_unit = 0
        
        # Общие затраты на единицу товара
        total_cost_per_unit = (
            cost_price +
            logistics_cnr_per_unit +
            logistics_rf_per_unit +
            duty_per_unit +
            conversion_fee_per_unit +
            credit_cost_per_unit
        )
        
        return {
            'cost_per_unit': total_cost_per_unit,
            'total_cost': total_cost_per_unit * quantity,
            'logistics_cnr_per_unit': logistics_cnr_per_unit,
            'logistics_rf_per_unit': logistics_rf_per_unit,
            'duty_per_unit': duty_per_unit,
            'conversion_fee_per_unit': conversion_fee_per_unit,
            'credit_cost_per_unit': credit_cost_per_unit
        }
    
    def calculate_positions(
        self,
        positions: List[Dict[str, Any]],
        logistics_rub: float,
        delivery_time: int,
        margin_percent: float,
        use_credit: bool = False,
        use_bank_guarantee: bool = False,
        payment_days: Optional[int] = None
    ) -> Dict[str, Any]:
        """Реализация интерфейса калькулятора."""
        return self.calculate_multi_position_prices(
            positions,
            logistics_rub,
            delivery_time,
            margin_percent,
            use_credit=use_credit,
            use_bank_guarantee=use_bank_guarantee,
            payment_days=payment_days
        )

    def calculate_multi_position_prices(
        self, positions: List[Dict[str, Any]],
        logistics_rub: float, delivery_time: int,
        target_margin_percent: float,
        use_credit: bool = False,
        use_bank_guarantee: bool = False,
        payment_days: Optional[int] = None,
        additional_expenses_total_yuan: float = 0
    ) -> Dict[str, Any]:
        """
        Рассчитывает цены для множественных позиций с единой итоговой маржой.
        
        Args:
            positions: Список позиций для расчета
            logistics_rub: Общая стоимость логистики (в рублях)
            delivery_time: Время доставки (в днях)
            target_margin_percent: Целевая маржа в процентах
            use_credit: Использовать ли кредит в расчете
            use_bank_guarantee: Использовать ли банковскую гарантию в расчете
            payment_days: Количество дней оплаты (для банковской гарантии)
            additional_expenses_total_yuan: Общая сумма дополнительных расходов (в юанях)
        
        Returns:
            Словарь с результатами расчета:
            - positions: Список позиций с рассчитанными ценами
            - total_costs: Общие затраты на все позиции
            - total_revenue: Общая выручка
            - target_margin: Целевая маржа
            - actual_margin: Фактическая маржа
            - price_coefficient: Коэффициент ценообразования (для режима 'global')
        """
        
        if not positions:
            return {
                'positions': [],
                'total_costs': 0,
                'total_revenue': 0,
                'target_margin': target_margin_percent,
                'actual_margin': 0,
                'price_coefficient': 1
            }
        
        # Рассчитываем общий вес всех позиций
        total_weight = sum(float(p['weight']) * int(p['quantity']) for p in positions)
        
        # Рассчитываем затраты для каждой позиции
        position_costs = []
        total_costs = 0
        
        for position in positions:
            costs = self.calculate_position_costs(
                position, logistics_rub, delivery_time, total_weight, 
                use_credit=use_credit, payment_days=payment_days
            )
            position_costs.append({
                'position': position,
                'costs': costs
            })
            total_costs += costs['total_cost']
        
        # Распределяем дополнительные расходы пропорционально затратам каждой позиции
        # Дополнительные расходы уже в юанях, все затраты тоже в юанях
        # Распределяем дополнительные расходы между позициями пропорционально их затратам
        if total_costs > 0 and additional_expenses_total_yuan > 0:
            for pos_data in position_costs:
                # Доля затрат позиции от общих затрат
                cost_ratio = pos_data['costs']['total_cost'] / total_costs
                # Дополнительные расходы для этой позиции (в юанях)
                position_additional_expenses_yuan = additional_expenses_total_yuan * cost_ratio
                # Добавляем к затратам на единицу (в юанях)
                pos_data['costs']['cost_per_unit'] += position_additional_expenses_yuan / int(pos_data['position']['quantity'])
                # Обновляем общие затраты позиции (в юанях)
                pos_data['costs']['total_cost'] += position_additional_expenses_yuan
        
        # Обновляем общие затраты с учетом дополнительных расходов (в юанях)
        total_costs += additional_expenses_total_yuan
        
        result_positions = []
        total_revenue = 0

        if self.pricing_mode == 'per_position':
            # Ценообразование: каждая позиция достигает целевой маржи
            # Если используется банковская гарантия, нужен итеративный расчет
            if use_bank_guarantee and payment_days is not None:
                # Итеративный расчет с учетом банковской гарантии для каждой позиции
                for pos_data in position_costs:
                    position = pos_data['position']
                    costs = pos_data['costs']
                    cost_per_unit = costs['cost_per_unit']
                    quantity = int(position['quantity'])
                    
                    # Начинаем с цены без банковской гарантии
                    if target_margin_percent >= 100:
                        final_price = cost_per_unit
                    else:
                        final_price = cost_per_unit / (1 - target_margin_percent / 100)
                    
                    # Итеративно уточняем цену с учетом банковской гарантии
                    # В Excel формула: I24*3%/365*(I15+I16), где I15+I16 = delivery_time + payment_days
                    bank_guarantee_days = delivery_time + payment_days  # K15 = I15 + I16
                    for _ in range(5):  # Максимум 5 итераций
                        revenue_with_vat = final_price * quantity * (1 + self.VAT_RATE)
                        bank_guarantee_cost = revenue_with_vat * 0.03 / 365 * bank_guarantee_days
                        bank_guarantee_cost_per_unit = bank_guarantee_cost / quantity if quantity > 0 else 0
                        
                        # Пересчитываем цену с учетом банковской гарантии
                        total_cost_with_guarantee = cost_per_unit + bank_guarantee_cost_per_unit
                        if target_margin_percent >= 100:
                            new_price = total_cost_with_guarantee
                        else:
                            new_price = total_cost_with_guarantee / (1 - target_margin_percent / 100)
                        
                        # Проверяем сходимость (разница менее 0.01)
                        if abs(new_price - final_price) < 0.01:
                            break
                        final_price = new_price
                    
                    general_price = final_price * quantity
                    total_revenue += general_price
                    result_positions.append({
                        'position': position,
                        'final_price': final_price,
                        'general_price': general_price,
                        'costs': costs,
                        'margin': (final_price - cost_per_unit) / final_price * 100 if final_price > 0 else 0
                    })
            else:
                # Обычный расчет без банковской гарантии
                for pos_data in position_costs:
                    position = pos_data['position']
                    costs = pos_data['costs']
                    cost_per_unit = costs['cost_per_unit']
                    if target_margin_percent >= 100:
                        final_price = cost_per_unit  # защита от деления на ноль; не повышаем цену
                    else:
                        final_price = cost_per_unit / (1 - target_margin_percent / 100)
                    quantity = int(position['quantity'])
                    general_price = final_price * quantity
                    total_revenue += general_price
                    result_positions.append({
                        'position': position,
                        'final_price': final_price,
                        'general_price': general_price,
                        'costs': costs,
                        'margin': (final_price - cost_per_unit) / final_price * 100 if final_price > 0 else 0
                    })
            price_coefficient = None
        else:
            # Ценообразование: общий коэффициент по целевой суммарной марже
            # Если используется банковская гарантия, нужен итеративный расчет
            if use_bank_guarantee and payment_days is not None:
                # Итеративный расчет с учетом банковской гарантии
                # В Excel формула: I24*3%/365*(I15+I16), где I15+I16 = delivery_time + payment_days
                # Банковская гарантия = (выручка с НДС) * 0.03 / 365 * (delivery_time + payment_days)
                bank_guarantee_days = delivery_time + payment_days  # K15 = I15 + I16
                target_revenue = total_costs / (1 - target_margin_percent / 100)
                price_coefficient = target_revenue / total_costs if total_costs > 0 else 1
                
                # Итеративно уточняем коэффициент с учетом банковской гарантии
                for _ in range(5):  # Максимум 5 итераций
                    # Рассчитываем выручку с текущим коэффициентом
                    test_revenue = sum(
                        costs['cost_per_unit'] * price_coefficient * int(pos_data['position']['quantity'])
                        for pos_data in position_costs
                    )
                    revenue_with_vat = test_revenue * (1 + self.VAT_RATE)
                    bank_guarantee_cost = revenue_with_vat * 0.03 / 365 * bank_guarantee_days
                    
                    # Пересчитываем коэффициент с учетом банковской гарантии
                    total_costs_with_guarantee = total_costs + bank_guarantee_cost
                    new_target_revenue = total_costs_with_guarantee / (1 - target_margin_percent / 100)
                    new_coefficient = new_target_revenue / total_costs if total_costs > 0 else 1
                    
                    # Проверяем сходимость (разница менее 0.0001)
                    if abs(new_coefficient - price_coefficient) < 0.0001:
                        break
                    price_coefficient = new_coefficient
            else:
                # Обычный расчет без банковской гарантии
                target_revenue = total_costs / (1 - target_margin_percent / 100)
                price_coefficient = target_revenue / total_costs if total_costs > 0 else 1
            
            for pos_data in position_costs:
                position = pos_data['position']
                costs = pos_data['costs']
                final_price = costs['cost_per_unit'] * price_coefficient
                quantity = int(position['quantity'])
                general_price = final_price * quantity
                total_revenue += general_price
                result_positions.append({
                    'position': position,
                    'final_price': final_price,
                    'general_price': general_price,
                    'costs': costs,
                    'margin': (final_price - costs['cost_per_unit']) / final_price * 100 if final_price > 0 else 0
                })
        
        # Проверяем итоговую маржу
        actual_margin = (total_revenue - total_costs) / total_revenue * 100 if total_revenue > 0 else 0
        
        return {
            'positions': result_positions,
            'total_costs': total_costs,
            'total_revenue': total_revenue,
            'target_margin': target_margin_percent,
            'actual_margin': actual_margin,
            'price_coefficient': price_coefficient
        }
    
    def calculate_legacy_single_position(
        self, position: Dict[str, Any], 
        logistics_rub: float, delivery_time: int, 
        margin_percent: float,
        use_credit: bool = False,
        use_bank_guarantee: bool = False,
        payment_days: Optional[int] = None,
        additional_expenses_total_yuan: float = 0
    ) -> Dict[str, Any]:
        """
        Рассчитывает цену для одной позиции (старый метод для совместимости).
        
        Использует старую логику расчета через calculate_selling_price.
        
        Args:
            position: Словарь с данными позиции
            logistics_rub: Общая стоимость логистики (в рублях)
            delivery_time: Время доставки (в днях)
            margin_percent: Целевая маржа в процентах
            use_credit: Использовать ли кредит в расчете
            use_bank_guarantee: Использовать ли банковскую гарантию в расчете
            payment_days: Количество дней оплаты (для банковской гарантии)
            additional_expenses_total_yuan: Общая сумма дополнительных расходов (в юанях)
        
        Returns:
            Словарь с результатами:
            - position: Исходные данные позиции
            - final_price: Финальная цена за единицу
            - general_price: Общая цена позиции
            - margin: Фактическая маржа
        """
        
        quantity = int(position['quantity'])
        cost_price = float(position['cost_price'])
        weight = float(position['weight'])
        duty_percent = float(position.get('duty_percent', 0))
        
        # Проверяем валидность данных
        if quantity <= 0 or weight <= 0:
            return {
                'position': position,
                'final_price': 0,
                'general_price': 0,
                'margin': 0
            }
        
        # Используем старую логику расчета
        final_price = calculate_selling_price(
            quantity=quantity,
            purchase_cost=cost_price,
            logistics_rub=logistics_rub,
            duty_percent=duty_percent,
            weight=weight,
            delivery_time=delivery_time,
            margin_percent=margin_percent,
            config=self.config,
            use_credit=use_credit,
            use_bank_guarantee=use_bank_guarantee,
            payment_days=payment_days,
            additional_expenses_total_yuan=additional_expenses_total_yuan
        )
        
        general_price = final_price * quantity
        
        return {
            'position': position,
            'final_price': final_price,
            'general_price': general_price,
            'margin': margin_percent
        }
