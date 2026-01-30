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
        additional_expenses_total_yuan: float = 0,
        position_margins: Optional[Dict[int, float]] = None
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
            position_margins: Словарь индивидуальных марж {position_index: margin_value}
        
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
        
        # Пересчитываем общие затраты как сумму всех позиций (уже включающих дополнительные расходы)
        # Это гарантирует, что дополнительные расходы учитываются только один раз
        total_costs = sum(pos_data['costs']['total_cost'] for pos_data in position_costs)
        
        # Разделяем позиции на фиксированные (с индивидуальной маржой) и пересчитываемые
        individual_margins = position_margins or {}
        fixed_position_indices = set(individual_margins.keys())
        variable_position_indices = set(range(len(position_costs))) - fixed_position_indices
        
        result_positions = []
        total_revenue = 0

        if self.pricing_mode == 'per_position':
            # Ценообразование: каждая позиция достигает целевой маржи
            # Если есть индивидуальные маржи, используем их для соответствующих позиций
            # ВАЖНО: все расчеты в юанях (как в Excel файле)
            # Если используется банковская гарантия, нужен итеративный расчет
            if use_bank_guarantee and payment_days is not None:
                # Точная формула на позицию: цена = cost / (1 - margin - (1+vat)*0.03/365*days)
                bank_guarantee_days = delivery_time + payment_days
                bg_rate = (1 + self.VAT_RATE) * 0.03 / 365 * bank_guarantee_days
                for idx, pos_data in enumerate(position_costs):
                    position = pos_data['position']
                    costs = pos_data['costs']
                    cost_per_unit = costs['cost_per_unit']
                    quantity = int(position['quantity'])
                    margin = individual_margins.get(idx, target_margin_percent)
                    margin_frac = margin / 100
                    if margin >= 100:
                        final_price = cost_per_unit
                    else:
                        denominator = (1 - margin_frac) - bg_rate
                        if denominator > 0:
                            final_price = cost_per_unit / denominator
                        else:
                            final_price = cost_per_unit / (1 - margin_frac)
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
                for idx, pos_data in enumerate(position_costs):
                    position = pos_data['position']
                    costs = pos_data['costs']
                    cost_per_unit = costs['cost_per_unit']
                    
                    # Определяем маржу для этой позиции
                    margin = individual_margins.get(idx, target_margin_percent)
                    
                    if margin >= 100:
                        final_price = cost_per_unit
                    else:
                        final_price = cost_per_unit / (1 - margin / 100)
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
            # ВАЖНО: все расчеты в юанях (как в Excel файле)
            
            # Если есть индивидуальные маржи, разделяем позиции на фиксированные и пересчитываемые
            if individual_margins:
                # Рассчитываем цены для фиксированных позиций (с индивидуальной маржой)
                fixed_revenue = 0
                fixed_costs = 0
                fixed_positions_results = []
                
                for idx in fixed_position_indices:
                    if idx < len(position_costs):
                        pos_data = position_costs[idx]
                        position = pos_data['position']
                        costs = pos_data['costs']
                        cost_per_unit = costs['cost_per_unit']
                        quantity = int(position['quantity'])
                        margin = individual_margins[idx]
                        
                        # Рассчитываем цену с индивидуальной маржой
                        if margin >= 100:
                            final_price = cost_per_unit
                        else:
                            final_price = cost_per_unit / (1 - margin / 100)
                        
                        general_price = final_price * quantity
                        fixed_revenue += general_price
                        fixed_costs += costs['total_cost']
                        
                        fixed_positions_results.append({
                            'index': idx,
                            'position': position,
                            'final_price': final_price,
                            'general_price': general_price,
                            'costs': costs,
                            'margin': margin
                        })
                
                # Рассчитываем оставшиеся затраты и целевую выручку для пересчитываемых позиций
                remaining_costs = total_costs - fixed_costs
                
                if remaining_costs <= 0:
                    # Все позиции имеют индивидуальную маржу, используем их как есть
                    for fixed_result in fixed_positions_results:
                        result_positions.append({
                            'position': fixed_result['position'],
                            'final_price': fixed_result['final_price'],
                            'general_price': fixed_result['general_price'],
                            'costs': fixed_result['costs'],
                            'margin': fixed_result['margin']
                        })
                    total_revenue = fixed_revenue
                    price_coefficient = None
                else:
                    # Рассчитываем коэффициент для пересчитываемых позиций
                    # Целевая общая выручка для достижения целевой маржи
                    total_target_revenue = total_costs / (1 - target_margin_percent / 100)
                    remaining_target_revenue = total_target_revenue - fixed_revenue
                    
                    # Проверяем возможность достижения целевой маржи
                    if remaining_target_revenue <= remaining_costs:
                        # Невозможно достичь целевой маржи - используем максимально возможную
                        # Выбрасываем предупреждение через коэффициент близкий к 1
                        price_coefficient = 1.0
                    else:
                        price_coefficient = remaining_target_revenue / remaining_costs if remaining_costs > 0 else 1
                    
                    # Если используется банковская гарантия — точная формула для целевой выручки
                    if use_bank_guarantee and payment_days is not None:
                        bank_guarantee_days = delivery_time + payment_days
                        margin_frac = target_margin_percent / 100
                        bg_rate = (1 + self.VAT_RATE) * 0.03 / 365 * bank_guarantee_days
                        denominator = (1 - margin_frac) - bg_rate
                        if denominator > 0:
                            total_target_revenue = total_costs / denominator
                            remaining_target_revenue = total_target_revenue - fixed_revenue
                            if remaining_target_revenue <= remaining_costs:
                                price_coefficient = 1.0
                            else:
                                price_coefficient = remaining_target_revenue / remaining_costs if remaining_costs > 0 else 1
                    
                    # Собираем результаты: сначала фиксированные позиции, потом пересчитываемые
                    result_dict = {res['index']: res for res in fixed_positions_results}
                    
                    for idx in variable_position_indices:
                        if idx < len(position_costs):
                            pos_data = position_costs[idx]
                            position = pos_data['position']
                            costs = pos_data['costs']
                            cost_per_unit = costs['cost_per_unit']
                            final_price = cost_per_unit * price_coefficient
                            quantity = int(position['quantity'])
                            general_price = final_price * quantity
                            
                            result_dict[idx] = {
                                'position': position,
                                'final_price': final_price,
                                'general_price': general_price,
                                'costs': costs,
                                'margin': (final_price - cost_per_unit) / final_price * 100 if final_price > 0 else 0
                            }
                    
                    # Сортируем результаты по индексу позиции
                    result_positions = [result_dict[i] for i in sorted(result_dict.keys())]
                    total_revenue = sum(pos['general_price'] for pos in result_positions)
            else:
                # Нет индивидуальных марж - используем стандартную логику
                if use_bank_guarantee and payment_days is not None:
                    # Точная формула: выручка без НДС R должна давать маржу target_margin с учётом БГ.
                    # R - total_costs - BG = target_margin * R,  BG = R * (1+vat) * 0.03/365 * days
                    # R * (1 - target_margin) - R * (1+vat)*0.03/365*days = total_costs
                    # R = total_costs / ((1 - target_margin) - (1+vat)*0.03/365*days)
                    bank_guarantee_days = delivery_time + payment_days
                    margin_frac = target_margin_percent / 100
                    bg_rate = (1 + self.VAT_RATE) * 0.03 / 365 * bank_guarantee_days
                    denominator = (1 - margin_frac) - bg_rate
                    if denominator > 0 and total_costs > 0:
                        target_revenue = total_costs / denominator
                        price_coefficient = target_revenue / total_costs
                    else:
                        target_revenue = total_costs / (1 - margin_frac)
                        price_coefficient = target_revenue / total_costs if total_costs > 0 else 1
                else:
                    # Обычный расчет без банковской гарантии
                    target_revenue = total_costs / (1 - target_margin_percent / 100)
                    price_coefficient = target_revenue / total_costs if total_costs > 0 else 1
                
                for pos_data in position_costs:
                    position = pos_data['position']
                    costs = pos_data['costs']
                    cost_per_unit = costs['cost_per_unit']
                    final_price = cost_per_unit * price_coefficient
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
        
        # Проверяем итоговую маржу (все в юанях). При банковской гарантии включаем её в затраты.
        total_costs_for_margin = total_costs
        if use_bank_guarantee and payment_days is not None and total_revenue > 0:
            bank_guarantee_days = delivery_time + payment_days
            revenue_with_vat = total_revenue * (1 + self.VAT_RATE)
            bank_guarantee_cost = revenue_with_vat * 0.03 / 365 * bank_guarantee_days
            total_costs_for_margin = total_costs + bank_guarantee_cost
        actual_margin = (total_revenue - total_costs_for_margin) / total_revenue * 100 if total_revenue > 0 else 0

        return {
            'positions': result_positions,
            'total_costs': total_costs,  # Возвращаем в юанях (как в Excel файле)
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
        # calculate_selling_price возвращает цену в юанях
        # ВАЖНО: final_price должен быть в юанях для Excel файла (затраты там тоже в юанях)
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
        
        # Рассчитываем затраты для расчета фактической маржи
        # Используем calculate_position_costs для получения базовых затрат
        total_weight = weight * quantity
        costs = self.calculate_position_costs(
            position, logistics_rub, delivery_time, total_weight,
            use_credit=use_credit, payment_days=payment_days
        )
        
        # Добавляем дополнительные расходы к затратам на единицу (в юанях)
        cost_per_unit_yuan = costs['cost_per_unit']
        if additional_expenses_total_yuan > 0:
            cost_per_unit_yuan += additional_expenses_total_yuan / quantity
        
        # Если используется банковская гарантия, нужно учесть её в затратах
        if use_bank_guarantee and payment_days is not None:
            # Рассчитываем стоимость банковской гарантии
            # В Excel формула: I24*3%/365*(I15+I16), где I15+I16 = delivery_time + payment_days
            # Выручка рассчитывается от цены в юанях
            bank_guarantee_days = delivery_time + payment_days
            revenue_with_vat = final_price * quantity * (1 + self.VAT_RATE)
            bank_guarantee_cost = revenue_with_vat * 0.03 / 365 * bank_guarantee_days
            bank_guarantee_cost_per_unit = bank_guarantee_cost / quantity if quantity > 0 else 0
            cost_per_unit_yuan += bank_guarantee_cost_per_unit
        
        # Рассчитываем фактическую маржу (все в юанях)
        actual_margin = (final_price - cost_per_unit_yuan) / final_price * 100 if final_price > 0 else 0
        
        general_price = final_price * quantity
        
        return {
            'position': position,
            'final_price': final_price,
            'general_price': general_price,
            'margin': actual_margin
        }
