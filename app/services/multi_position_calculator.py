# app/services/multi_position_calculator.py
from typing import List, Dict, Any, Tuple
from app.calculate import calculate_selling_price


class MultiPositionCalculator:
    """Калькулятор для множественных позиций с единой итоговой маржой"""
    
    def __init__(self, config=None):
        self.config = config or {}
        self.calc_config = self.config.get('calculation_constants', {})
        
        # Константы расчета
        self.CONVERSION_RATE = self.calc_config.get('conversion_rate', 12)
        self.LOGISTICS_CNR_RATIO = self.calc_config.get('logistics_cnr_ratio', 0.3)
        self.LOGISTICS_RF_RATIO = self.calc_config.get('logistics_rf_ratio', 0.7)
        self.CONVERSION_FEE_RATE = self.calc_config.get('conversion_fee_rate', 0.032)
        self.CREDIT_RATE = self.calc_config.get('credit_rate', 0.16)
    
    def calculate_position_costs(self, position: Dict[str, Any], logistics_rub: float, 
                               delivery_time: int, total_weight: float) -> Dict[str, float]:
        """Рассчитывает все затраты для одной позиции"""
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
        
        # Расчет кредитных затрат
        credit_cost = cost_price * quantity * self.CREDIT_RATE / 365 * delivery_time
        credit_cost_per_unit = credit_cost / quantity if quantity > 0 else 0
        
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
    
    def calculate_multi_position_prices(self, positions: List[Dict[str, Any]], 
                                      logistics_rub: float, delivery_time: int, 
                                      target_margin_percent: float) -> Dict[str, Any]:
        """Рассчитывает цены для множественных позиций с единой итоговой маржой"""
        
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
            costs = self.calculate_position_costs(position, logistics_rub, delivery_time, total_weight)
            position_costs.append({
                'position': position,
                'costs': costs
            })
            total_costs += costs['total_cost']
        
        # Рассчитываем коэффициент для достижения целевой маржи
        # total_revenue = total_costs / (1 - target_margin_percent / 100)
        # coefficient = total_revenue / total_costs
        target_revenue = total_costs / (1 - target_margin_percent / 100)
        price_coefficient = target_revenue / total_costs if total_costs > 0 else 1
        
        # Рассчитываем цены для каждой позиции
        result_positions = []
        total_revenue = 0
        
        for pos_data in position_costs:
            position = pos_data['position']
            costs = pos_data['costs']
            
            # Применяем коэффициент к затратам для получения цены
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
    
    def calculate_legacy_single_position(self, position: Dict[str, Any], 
                                       logistics_rub: float, delivery_time: int, 
                                       margin_percent: float) -> Dict[str, Any]:
        """Рассчитывает цену для одной позиции (старый метод для совместимости)"""
        
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
            config=self.config
        )
        
        general_price = final_price * quantity
        
        return {
            'position': position,
            'final_price': final_price,
            'general_price': general_price,
            'margin': margin_percent
        }
