"""
Модуль расчета логистики по правилам ООО ТД РИНАКО.

Основные правила:
- Стандартная еврофура: 13600×2450×2650 мм, 20 тонн, 88 м³
- Трал 40 тонн: для тяжеловесных и негабаритных грузов
- Расчет по весу или объему (выбирается превалирующий параметр)
- Для грузов >5 тонн используются цены основных городов
- Для грузов до 600кг рекомендуется Деловые линии
"""

import math
from typing import Dict, List, Optional, Tuple


# Стандартные параметры транспорта
EURO_TRUCK_LENGTH_MM = 13600
EURO_TRUCK_WIDTH_MM = 2450
EURO_TRUCK_HEIGHT_MM = 2650
EURO_TRUCK_CAPACITY_KG = 20000
EURO_TRUCK_VOLUME_M3 = 88

TRAIL_CAPACITY_KG = 40000

# Пороговые значения
MIN_WEIGHT_FOR_MAIN_CITIES_KG = 5000
SMALL_CARGO_THRESHOLD_KG = 600
SMALL_CARGO_RECOMMENDATION_KG = 700


def calculate_cargo_volume(length_mm: float, width_mm: float, height_mm: float) -> float:
    """Вычисляет объем груза в м³."""
    return (length_mm * width_mm * height_mm) / 1_000_000_000


def is_oversized(length_mm: float, width_mm: float, height_mm: float) -> bool:
    """Проверяет, является ли груз негабаритным (превышает размеры еврофуры)."""
    return (
        length_mm > EURO_TRUCK_LENGTH_MM
        or width_mm > EURO_TRUCK_WIDTH_MM
        or height_mm > EURO_TRUCK_HEIGHT_MM
    )


def is_heavy(weight_kg: float) -> bool:
    """Проверяет, является ли груз тяжеловесным (превышает грузоподъемность фуры)."""
    return weight_kg > EURO_TRUCK_CAPACITY_KG


def determine_calculation_basis(
    weight_kg: float,
    volume_m3: float,
    truck_capacity_kg: int = EURO_TRUCK_CAPACITY_KG,
    truck_volume_m3: float = EURO_TRUCK_VOLUME_M3
) -> str:
    """
    Определяет, по какому параметру рассчитывать логистику (вес или объем).
    
    Возвращает 'weight' если вес превалирует, 'volume' если объем превалирует.
    """
    weight_ratio = weight_kg / truck_capacity_kg
    volume_ratio = volume_m3 / truck_volume_m3
    
    # Если один из параметров превышает грузоподъемность/объем, используем его
    if weight_ratio > 1.0 and volume_ratio <= 1.0:
        return 'weight'
    if volume_ratio > 1.0 and weight_ratio <= 1.0:
        return 'volume'
    
    # Выбираем превалирующий параметр
    return 'weight' if weight_ratio >= volume_ratio else 'volume'


def calculate_logistics_by_weight(
    weight_kg: float,
    city_price: float,
    truck_capacity_kg: int = EURO_TRUCK_CAPACITY_KG
) -> Dict[str, any]:
    """
    Рассчитывает логистику по весу.
    
    Формула: (Стоимость до города / Грузоподъемность) × Вес груза
    """
    price_per_kg = city_price / truck_capacity_kg
    total_price = price_per_kg * weight_kg
    
    trucks_count = (weight_kg + truck_capacity_kg - 1) // truck_capacity_kg  # Округление вверх
    
    return {
        'basis': 'weight',
        'price_per_kg': price_per_kg,
        'total_price': total_price,
        'trucks_count': trucks_count,
        'calculation_formula': f'({city_price:,.0f} руб / {truck_capacity_kg:,} кг) × {weight_kg:,.0f} кг',
    }


def calculate_logistics_by_volume(
    volume_m3: float,
    city_price: float,
    truck_volume_m3: float = EURO_TRUCK_VOLUME_M3
) -> Dict[str, any]:
    """
    Рассчитывает логистику по объему.
    
    Формула: (Стоимость фуры / Полный полезный объем кузова) × Объем груза
    """
    price_per_m3 = city_price / truck_volume_m3
    total_price = price_per_m3 * volume_m3
    
    # Округление вверх для количества машин
    trucks_count = math.ceil(volume_m3 / truck_volume_m3)
    
    return {
        'basis': 'volume',
        'price_per_m3': price_per_m3,
        'total_price': total_price,
        'trucks_count': trucks_count,
        'calculation_formula': f'({city_price:,.0f} руб / {truck_volume_m3:.0f} м³) × {volume_m3:.2f} м³',
    }


def calculate_logistics(
    weight_kg: float,
    city_price: float,
    transport_type: str = 'truck',
    length_mm: Optional[float] = None,
    width_mm: Optional[float] = None,
    height_mm: Optional[float] = None,
) -> Dict[str, any]:
    """
    Основная функция расчета логистики.
    
    Args:
        weight_kg: Вес груза в кг
        city_price: Стоимость доставки полной фуры/трала до города
        transport_type: 'truck' (фура) или 'trail' (трал)
        length_mm: Длина груза в мм (опционально)
        width_mm: Ширина груза в мм (опционально)
        height_mm: Высота груза в мм (опционально)
    
    Returns:
        Словарь с результатами расчета
    """
    truck_capacity = TRAIL_CAPACITY_KG if transport_type == 'trail' else EURO_TRUCK_CAPACITY_KG
    truck_volume = EURO_TRUCK_VOLUME_M3  # Объем одинаковый для фуры и трала
    
    # Для мелкогабаритных грузов
    if weight_kg < SMALL_CARGO_THRESHOLD_KG:
        return {
            'basis': 'small_cargo',
            'total_price': None,
            'recommendation': 'dellin',
            'message': f'Для грузов менее {SMALL_CARGO_THRESHOLD_KG} кг рекомендуется транспортная компания «Деловые линии» (dellin.ru). Логистика по Китаю + 30% к стоимости из просчета.',
        }
    
    # Если габариты не указаны (None), считаем только по весу
    if length_mm is None or width_mm is None or height_mm is None:
        return calculate_logistics_by_weight(weight_kg, city_price, truck_capacity)
    
    # Вычисляем объем
    volume_m3 = calculate_cargo_volume(length_mm, width_mm, height_mm)
    
    # Определяем, является ли груз негабаритным или тяжеловесным
    oversized = is_oversized(length_mm, width_mm, height_mm)
    heavy = is_heavy(weight_kg)
    
    # Для негабаритных или тяжеловесных грузов может потребоваться трал
    if oversized or heavy:
        if transport_type == 'truck' and heavy:
            # Тяжеловесный груз требует трал
            truck_capacity = TRAIL_CAPACITY_KG
        elif transport_type == 'truck' and oversized:
            # Негабаритный груз может потребовать трал, но оставляем выбор пользователю
            pass
    
    # Определяем базис расчета (вес или объем)
    basis = determine_calculation_basis(weight_kg, volume_m3, truck_capacity, truck_volume)
    
    if basis == 'weight':
        result = calculate_logistics_by_weight(weight_kg, city_price, truck_capacity)
    else:
        result = calculate_logistics_by_volume(volume_m3, city_price, truck_volume)
    
    # Добавляем информацию о габаритах
    result.update({
        'weight_kg': weight_kg,
        'volume_m3': volume_m3,
        'dimensions': {
            'length_mm': length_mm,
            'width_mm': width_mm,
            'height_mm': height_mm,
        },
        'oversized': oversized,
        'heavy': heavy,
        'transport_type': transport_type,
    })
    
    return result


