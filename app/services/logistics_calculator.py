"""
Модуль расчета логистики по правилам ООО ТД РИНАКО.

Основные правила:
- Стандартная еврофура: 13600×2450×2650 мм, 20 тонн, 82 м³
- Трал 40 тонн: для тяжеловесных и негабаритных грузов (только до Екатеринбурга и Москвы)
- Расчет по весу или объему (выбирается превалирующий параметр)
- Для грузов >5 тонн используются цены основных городов
- Для грузов до 600кг рекомендуется Деловые линии
- Для грузов <5 тонн в города вне основного маршрута: алгоритм "До ЕКБ + по РФ"
"""

import logging
import math
from typing import Any, Dict, Optional, Tuple

logger = logging.getLogger(__name__)


# Стандартные параметры транспорта
EURO_TRUCK_LENGTH_MM = 13600
EURO_TRUCK_WIDTH_MM = 2450
EURO_TRUCK_HEIGHT_MM = 2650
EURO_TRUCK_CAPACITY_KG = 20000
EURO_TRUCK_VOLUME_M3 = 82  # Обновлено согласно новым правилам

TRAIL_CAPACITY_KG = 40000
TRAIL_LENGTH_MM = 16000  # Максимальная длина трала
TRAIL_WIDTH_MM = 3500  # Общая ширина: 2500 мм + 500 мм по бокам
TRAIL_HEIGHT_MM = 4500  # Высота трала
TRAIL_VOLUME_M3 = 250  # Объем трала
TRAIL_REAR_SPACE_MM = 2000  # Дополнительное пространство сзади

# Пороговые значения
MIN_WEIGHT_FOR_MAIN_CITIES_KG = 5000
WEIGHT_THRESHOLD_FOR_300KM_RULE_KG = 5000  # 5 тонн для применения правила 300км
SMALL_CARGO_THRESHOLD_KG = 600
SMALL_CARGO_RECOMMENDATION_KG = 800  # Обновлено с 700 на 800 кг

# Города, куда доступен трал
TRAIL_AVAILABLE_CITIES = ['Екатеринбург', 'Москва']

# Тарифная сетка для перевозки по РФ от Екатеринбурга (руб за 1000 км)
RF_TARIFF_TABLE = {
    1000: 30000,    # 1 тонна за 1000 км
    2000: 52000,    # 2 тонны
    3000: 69000,    # 3 тонны
    5000: 90000,    # 5 тонн
    7000: 105000,   # 7 тонн
    10000: 115000,  # 10 тонн
    15000: 118000,  # 15 тонн
    20000: 120000   # 20 тонн (полная фура)
}

# Стоимость фуры до Екатеринбурга (для расчета по алгоритму "До ЕКБ + по РФ")
EKB_TRUCK_PRICE = 1100000


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


def fits_in_one_truck(
    weight_kg: float,
    length_mm: Optional[float] = None,
    width_mm: Optional[float] = None,
    height_mm: Optional[float] = None
) -> bool:
    """
    Проверяет, помещается ли груз в одну фуру.
    
    Args:
        weight_kg: Вес груза в кг
        length_mm: Длина груза в мм (опционально)
        width_mm: Ширина груза в мм (опционально)
        height_mm: Высота груза в мм (опционально)
    
    Returns:
        True если груз помещается в одну фуру, False если не помещается
    """
    # Проверка по весу
    if weight_kg > EURO_TRUCK_CAPACITY_KG:
        return False
    
    # Если габариты не указаны, проверяем только по весу
    if length_mm is None or width_mm is None or height_mm is None:
        return weight_kg <= EURO_TRUCK_CAPACITY_KG
    
    # Проверка габаритов
    if is_oversized(length_mm, width_mm, height_mm):
        return False
    
    # Проверка по объему
    volume_m3 = calculate_cargo_volume(length_mm, width_mm, height_mm)
    if volume_m3 > EURO_TRUCK_VOLUME_M3:
        return False
    
    return True


def is_oversized_for_trail(length_mm: float, width_mm: float, height_mm: float) -> Tuple[bool, Optional[str]]:
    """
    Проверяет, является ли груз негабаритным для трала.
    
    Args:
        length_mm: Длина груза в мм
        width_mm: Ширина груза в мм
        height_mm: Высота груза в мм
    
    Returns:
        Кортеж (is_oversized, error_message):
        - is_oversized: True если груз не помещается в трал, False если помещается
        - error_message: Сообщение об ошибке с указанием конкретного параметра, если превышен
    """
    issues = []
    
    # Проверка габаритов
    if length_mm > TRAIL_LENGTH_MM:
        issues.append(f'длина ({length_mm:,.0f} мм превышает максимальную {TRAIL_LENGTH_MM:,.0f} мм)')
    
    if width_mm > TRAIL_WIDTH_MM:
        issues.append(f'ширина ({width_mm:,.0f} мм превышает максимальную {TRAIL_WIDTH_MM:,.0f} мм)')
    
    if height_mm > TRAIL_HEIGHT_MM:
        issues.append(f'высота ({height_mm:,.0f} мм превышает максимальную {TRAIL_HEIGHT_MM:,.0f} мм)')
    
    # Проверка объема
    volume_m3 = calculate_cargo_volume(length_mm, width_mm, height_mm)
    if volume_m3 > TRAIL_VOLUME_M3:
        issues.append(f'объем ({volume_m3:.2f} м³ превышает максимальный {TRAIL_VOLUME_M3} м³)')
    
    if issues:
        error_message = 'Груз превышает параметры трала по ' + ', '.join(issues) + '. Пожалуйста, обратитесь к логисту для расчета индивидуальной стоимости.'
        return True, error_message
    
    return False, None


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
) -> Dict[str, Any]:
    """
    Рассчитывает логистику по весу.
    
    Формула: (Стоимость до города / Грузоподъемность) × Вес груза
    """
    import logging
    logger = logging.getLogger(__name__)
    
    # Логируем входные параметры для отладки
    logger.debug(f"calculate_logistics_by_weight: weight_kg={weight_kg}, city_price={city_price}, truck_capacity_kg={truck_capacity_kg}")
    
    # Проверка на разумность веса (максимум 1000 тонн = 1,000,000 кг)
    if weight_kg > 1_000_000:
        logger.warning(f"Подозрительно большой вес в calculate_logistics_by_weight: {weight_kg} кг. Возможно ошибка.")
    
    price_per_kg = city_price / truck_capacity_kg
    total_price = price_per_kg * weight_kg
    
    trucks_count = (weight_kg + truck_capacity_kg - 1) // truck_capacity_kg  # Округление вверх
    
    result = {
        'basis': 'weight',
        'weight_kg': weight_kg,  # Добавляем вес для отображения
        'price_per_kg': price_per_kg,
        'total_price': total_price,
        'trucks_count': trucks_count,
        'calculation_formula': f'({city_price:,.0f} руб / {truck_capacity_kg:,} кг) × {weight_kg:,.0f} кг',
    }
    
    logger.debug(f"calculate_logistics_by_weight результат: weight_kg={result['weight_kg']}, formula={result['calculation_formula']}")
    
    return result


def calculate_logistics_by_volume(
    volume_m3: float,
    city_price: float,
    truck_volume_m3: float = EURO_TRUCK_VOLUME_M3
) -> Dict[str, Any]:
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


def get_rf_tariff_per_1000km(weight_kg: float) -> float:
    """
    Возвращает тариф за 1000 км для указанного веса по таблице РФ.
    
    Для промежуточных весов возвращается ставка для ближайшего большего веса.
    """
    if weight_kg <= 0:
        return 0.0
    
    # Находим ближайший больший или равный вес в таблице
    sorted_weights = sorted(RF_TARIFF_TABLE.keys())
    
    for table_weight in sorted_weights:
        if weight_kg <= table_weight:
            return RF_TARIFF_TABLE[table_weight]
    
    # Если вес превышает максимальный в таблице, используем ставку для 20 тонн
    return RF_TARIFF_TABLE[20000]


def calculate_ekb_plus_rf_route(
    weight_kg: float,
    distance_from_ekb_km: int,
    truck_capacity_kg: int = EURO_TRUCK_CAPACITY_KG
) -> Dict[str, Any]:
    """
    Рассчитывает стоимость по алгоритму: КНР → ЕКБ + ЕКБ → Город.
    
    Используется для грузов в города из справочника ЕКБ+РФ.
    Расчет от ЕКБ до города выполняется отдельно для каждой фуры.
    
    Args:
        weight_kg: Вес груза в кг
        distance_from_ekb_km: Расстояние от Екатеринбурга до города назначения в км
        truck_capacity_kg: Грузоподъемность транспорта
    
    Returns:
        Словарь с результатами расчета
    """
    # Шаг 1: Рассчитать стоимость от КНР до ЕКБ
    china_to_ekb_result = calculate_logistics_by_weight(weight_kg, EKB_TRUCK_PRICE, truck_capacity_kg)
    china_to_ekb_price = china_to_ekb_result['total_price']
    trucks_count = int(china_to_ekb_result['trucks_count'])  # Преобразуем в int для range()
    
    # Шаг 2: Рассчитать стоимость от ЕКБ до города назначения по РФ
    # Для каждой фуры рассчитываем отдельно
    remaining_weight = weight_kg
    ekb_to_destination_price = 0.0
    truck_details = []
    
    # Преобразуем distance_from_ekb_km в int для форматирования
    distance_km_int = int(distance_from_ekb_km)
    
    for truck_num in range(trucks_count):
        # Определяем вес для текущей фуры
        if remaining_weight >= truck_capacity_kg:
            truck_weight = truck_capacity_kg
            remaining_weight -= truck_capacity_kg
        else:
            truck_weight = remaining_weight
            remaining_weight = 0
        
        # Получаем тариф для веса этой фуры
        tariff_per_1000km = get_rf_tariff_per_1000km(truck_weight)
        truck_price = (tariff_per_1000km / 1000) * distance_from_ekb_km
        ekb_to_destination_price += truck_price
        
        truck_details.append({
            'truck_number': truck_num + 1,
            'weight_kg': truck_weight,
            'tariff_per_1000km': tariff_per_1000km,
            'price': truck_price,
            'formula': f'({tariff_per_1000km:,.0f} руб / 1000 км) × {distance_km_int:,} км'
        })
    
    # Итоговая стоимость
    total_price = china_to_ekb_price + ekb_to_destination_price
    
    # Формируем формулу для отображения
    if len(truck_details) == 1:
        ekb_formula = truck_details[0]['formula']
        # Для одной фуры используем её тариф
        ekb_tariff_per_1000km = truck_details[0]['tariff_per_1000km']
    else:
        ekb_formula = ' + '.join([f'Фура {td["truck_number"]}: {td["formula"]}' for td in truck_details])
        # Для нескольких фур используем тариф для общего веса груза
        ekb_tariff_per_1000km = get_rf_tariff_per_1000km(weight_kg)
    
    return {
        'basis': 'weight',
        'weight_kg': weight_kg,  # Добавляем вес для отображения
        'calculation_type': 'ekb_plus_rf',
        'price_per_kg': china_to_ekb_result['price_per_kg'],
        'total_price': total_price,
        'trucks_count': trucks_count,
        'route_details': {
            'china_to_ekb': {
                'price': china_to_ekb_price,
                'formula': china_to_ekb_result['calculation_formula'],
            },
            'ekb_to_destination': {
                'price': ekb_to_destination_price,
                'distance_km': distance_km_int,
                'tariff_per_1000km': ekb_tariff_per_1000km,
                'trucks': truck_details,
                'formula': ekb_formula,
            }
        },
        'calculation_formula': f'КНР→ЕКБ: {china_to_ekb_price:,.0f} руб + ЕКБ→Город: {ekb_to_destination_price:,.0f} руб',
    }


def calculate_logistics(
    weight_kg: float,
    city_price: float,
    transport_type: str = 'truck',
    city_name: Optional[str] = None,
    is_main_route: bool = True,
    distance_from_ekb_km: Optional[int] = None,
    length_mm: Optional[float] = None,
    width_mm: Optional[float] = None,
    height_mm: Optional[float] = None,
    city_in_ekb_rf_catalog: Optional[bool] = None,
) -> Dict[str, Any]:
    """
    Основная функция расчета логистики.
    
    Args:
        weight_kg: Вес груза в кг
        city_price: Стоимость доставки полной фуры/трала до города
        transport_type: 'truck' (фура) или 'trail' (трал)
        city_name: Название города назначения (для проверки доступности трала)
        is_main_route: Город находится на основном маршруте или в радиусе 300км
        distance_from_ekb_km: Расстояние от Екатеринбурга (для алгоритма "До ЕКБ + по РФ")
        length_mm: Длина груза в мм (опционально)
        width_mm: Ширина груза в мм (опционально)
        height_mm: Высота груза в мм (опционально)
        city_in_ekb_rf_catalog: Явное указание, что город в справочнике ЕКБ+РФ (если None, проверяется автоматически)
    
    Returns:
        Словарь с результатами расчета
    """
    truck_capacity = TRAIL_CAPACITY_KG if transport_type == 'trail' else EURO_TRUCK_CAPACITY_KG
    truck_volume = TRAIL_VOLUME_M3 if transport_type == 'trail' else EURO_TRUCK_VOLUME_M3
    
    # Для мелкогабаритных грузов (≤ 600 кг)
    if weight_kg <= SMALL_CARGO_THRESHOLD_KG:
        return {
            'basis': 'small_cargo',
            'weight_kg': weight_kg,  # Добавляем вес для отображения
            'total_price': None,
            'recommendation': 'dellin',
            'message': f'Для грузов до {SMALL_CARGO_THRESHOLD_KG} кг рекомендуется транспортная компания «Деловые линии» (dellin.ru). Расчет от Забайкальска (погранпереход) до точки назначения. Логистика по Китаю + 30% к стоимости из просчета.',
        }
    
    # Проверка доступности трала для города (только если трал выбран явно, не при автопереключении)
    # Автопереключение проверяется позже, когда известны габариты груза
    if transport_type == 'trail' and (length_mm is None or width_mm is None or height_mm is None):
        # Если трал выбран явно без габаритов, проверяем доступность
        if city_name and city_name not in TRAIL_AVAILABLE_CITIES:
            return {
                'error': f'Трал доступен только до городов: {", ".join(TRAIL_AVAILABLE_CITIES)}. Для города {city_name} трал недоступен.',
                'trail_not_available': True,
            }
    
    # Проверяем негабаритность груза ДО выбора алгоритма расчета
    # Логика: если груз не помещается в одну фуру и негабаритный - это один большой груз
    # Для городов вне основного маршрута (не ЕКБ/Москва) выдаем ошибку
    if length_mm is not None and width_mm is not None and height_mm is not None:
        oversized = is_oversized(length_mm, width_mm, height_mm)
        fits_in_one = fits_in_one_truck(weight_kg, length_mm, width_mm, height_mm)
        
        # Если груз не помещается в одну фуру И негабаритный - это один большой груз
        if not fits_in_one and oversized and transport_type == 'truck':
            # Если груз негабаритный для фуры
            if city_name and city_name in TRAIL_AVAILABLE_CITIES:
                # Для ЕКБ и Москвы переключаемся на трал (проверка будет позже)
                pass
            elif not is_main_route:
                # Для городов вне основного маршрута (не ЕКБ/Москва) - ошибка
                return {
                    'error': 'Негабаритный груз для городов вне основного маршрута (кроме Екатеринбурга и Москвы) требует индивидуального расчета. Пожалуйста, обратитесь к логисту.',
                    'oversized': True,
                    'weight_kg': weight_kg,
                    'dimensions': {
                        'length_mm': length_mm,
                        'width_mm': width_mm,
                        'height_mm': height_mm,
                    },
                }
    
    # Проверяем, находится ли город в справочнике ЕКБ+РФ
    from app.services.datasets import get_ekb_rf_city_distance, is_city_in_ekb_rf_catalog
    
    if city_in_ekb_rf_catalog is None:
        # Автоматическая проверка, если не указано явно
        city_in_ekb_rf_catalog = city_name is not None and is_city_in_ekb_rf_catalog(city_name)
    
    # Если город в справочнике ЕКБ+РФ, получаем расстояние из справочника если не указано
    if city_in_ekb_rf_catalog and distance_from_ekb_km is None and city_name:
        distance_from_ekb_km = get_ekb_rf_city_distance(city_name)
    
    # Определяем, нужен ли расчет по алгоритму "До ЕКБ + по РФ"
    # Условия:
    # 1. Город в справочнике ЕКБ+РФ ИЛИ
    # 2. (вес < 5 тонн И не основной маршрут) И есть расстояние от ЕКБ
    use_ekb_plus_rf = False
    if city_in_ekb_rf_catalog:
        # Для городов из справочника ЕКБ+РФ всегда используем алгоритм, независимо от веса
        if distance_from_ekb_km is not None and distance_from_ekb_km > 0:
            use_ekb_plus_rf = True
        else:
            # Если расстояние не указано, запрашиваем его
            return {
                'error': f'Для города {city_name} из справочника ЕКБ+РФ необходимо указать расстояние от Екатеринбурга.',
                'needs_distance': True,
                'weight_kg': weight_kg,
                'city_in_ekb_rf_catalog': True,
            }
    elif (
        weight_kg < WEIGHT_THRESHOLD_FOR_300KM_RULE_KG
        and not is_main_route
        and distance_from_ekb_km is not None
        and distance_from_ekb_km > 0
    ):
        # Для городов вне справочника ЕКБ+РФ используем старую логику (только для грузов < 5 тонн)
        use_ekb_plus_rf = True
    
    if use_ekb_plus_rf:
        # Используем алгоритм "До ЕКБ + по РФ"
        # Проверяем негабаритность для алгоритма ЕКБ+РФ
        # Логика: если груз не помещается в одну фуру и негабаритный - это один большой груз
        if length_mm is not None and width_mm is not None and height_mm is not None:
            oversized = is_oversized(length_mm, width_mm, height_mm)
            fits_in_one = fits_in_one_truck(weight_kg, length_mm, width_mm, height_mm)
            
            # Если груз не помещается в одну фуру И негабаритный - это один большой груз
            if not fits_in_one and oversized and transport_type == 'truck':
                # Если груз негабаритный для фуры
                if city_name and city_name in TRAIL_AVAILABLE_CITIES:
                    # Для ЕКБ и Москвы переключаемся на трал
                    transport_type = 'trail'
                    truck_capacity = TRAIL_CAPACITY_KG
                    # Обновляем city_price для трала
                    if city_name:
                        from app.services.datasets import load_trail_cities
                        trail_cities = load_trail_cities()
                        for city in trail_cities:
                            if city.get('name') == city_name:
                                trail_price = city.get('trail_price')
                                if trail_price and trail_price > 0:
                                    city_price = trail_price
                                break
                elif not is_main_route:
                    # Для городов вне основного маршрута (не ЕКБ/Москва) - ошибка
                    return {
                        'error': 'Негабаритный груз для городов вне основного маршрута (кроме Екатеринбурга и Москвы) требует индивидуального расчета. Пожалуйста, обратитесь к логисту.',
                        'oversized': True,
                        'weight_kg': weight_kg,
                        'dimensions': {
                            'length_mm': length_mm,
                            'width_mm': width_mm,
                            'height_mm': height_mm,
                        },
                    }
        
        # Преобразуем distance_from_ekb_km в int если он float
        distance_int = int(distance_from_ekb_km) if distance_from_ekb_km is not None else None
        result = calculate_ekb_plus_rf_route(weight_kg, distance_int, truck_capacity)
        result['weight_kg'] = weight_kg
        result['transport_type'] = transport_type
        result['distance_from_ekb_provided'] = True
        result['can_save_distance'] = False  # Расстояние уже сохранено
        result['city_in_ekb_rf_catalog'] = city_in_ekb_rf_catalog
        return result
    
    # Проверка: если груз < 5 тонн, не основной маршрут и расстояние не указано
    if weight_kg < WEIGHT_THRESHOLD_FOR_300KM_RULE_KG and not is_main_route and distance_from_ekb_km is None:
        return {
            'error': 'Для грузов менее 5 тонн в города вне основного маршрута необходимо указать расстояние от Екатеринбурга.',
            'needs_distance': True,
            'weight_kg': weight_kg,
        }
    
    # Стандартный расчет (прямой расчет по формуле 5.2)
    # Если габариты не указаны (None), считаем только по весу
    if length_mm is None or width_mm is None or height_mm is None:
        logger.debug(f"calculate_logistics: расчет по весу, weight_kg={weight_kg}, city_price={city_price}, truck_capacity={truck_capacity}")
        result = calculate_logistics_by_weight(weight_kg, city_price, truck_capacity)
        # Убеждаемся, что weight_kg правильный (защита от перезаписи)
        if result.get('weight_kg') != weight_kg:
            logger.warning(f"calculate_logistics: weight_kg в результате ({result.get('weight_kg')}) не совпадает с переданным ({weight_kg}), исправляем")
            result['weight_kg'] = weight_kg
            # Пересчитываем формулу с правильным весом
            result['calculation_formula'] = f'({city_price:,.0f} руб / {truck_capacity:,} кг) × {weight_kg:,.0f} кг'
        result['calculation_type'] = 'direct'
        result['transport_type'] = transport_type
        result['distance_from_ekb_provided'] = False
        result['can_save_distance'] = distance_from_ekb_km is None and not is_main_route
        logger.debug(f"calculate_logistics: результат weight_kg={result.get('weight_kg')}, formula={result.get('calculation_formula')}")
        return result
    
    # Вычисляем объем
    volume_m3 = calculate_cargo_volume(length_mm, width_mm, height_mm)
    
    # Определяем, является ли груз негабаритным или тяжеловесным
    oversized = is_oversized(length_mm, width_mm, height_mm)
    heavy = is_heavy(weight_kg)
    
    # Автоматическое переключение на трал для негабаритных грузов
    auto_switched_to_trail = False
    if transport_type == 'truck' and oversized and length_mm is not None and width_mm is not None and height_mm is not None:
        # Проверяем доступность трала для города
        if city_name and city_name in TRAIL_AVAILABLE_CITIES:
            # Проверяем, помещается ли груз в трал
            is_oversized_trail, trail_error_message = is_oversized_for_trail(length_mm, width_mm, height_mm)
            if is_oversized_trail:
                # Груз не помещается в трал - нужно обратиться к логисту
                return {
                    'error': trail_error_message or 'Груз превышает габариты или объем трала. Пожалуйста, обратитесь к логисту для расчета индивидуальной стоимости.',
                    'exceeds_trail': True,
                    'weight_kg': weight_kg,
                    'dimensions': {
                        'length_mm': length_mm,
                        'width_mm': width_mm,
                        'height_mm': height_mm,
                    },
                    'volume_m3': volume_m3,
                }
            else:
                # Груз помещается в трал - автоматически переключаемся
                transport_type = 'trail'
                truck_capacity = TRAIL_CAPACITY_KG
                truck_volume = TRAIL_VOLUME_M3
                auto_switched_to_trail = True
                # Обновляем city_price для трала, если доступен
                if city_name:
                    from app.services.datasets import load_trail_cities
                    trail_cities = load_trail_cities()
                    for city in trail_cities:
                        if city.get('name') == city_name:
                            trail_price = city.get('trail_price')
                            if trail_price and trail_price > 0:
                                city_price = trail_price
                            break
        elif city_name and city_name not in TRAIL_AVAILABLE_CITIES:
            # Трал недоступен для этого города
            return {
                'error': f'Груз не помещается в фуру, но трал доступен только до городов: {", ".join(TRAIL_AVAILABLE_CITIES)}. Для города {city_name} трал недоступен.',
                'trail_not_available': True,
                'oversized': True,
                'weight_kg': weight_kg,
            }
    
    # Для тяжеловесных грузов также может потребоваться трал
    if transport_type == 'truck' and heavy:
        # Тяжеловесный груз требует трал
        if city_name and city_name in TRAIL_AVAILABLE_CITIES:
            transport_type = 'trail'
            truck_capacity = TRAIL_CAPACITY_KG
            truck_volume = TRAIL_VOLUME_M3
            auto_switched_to_trail = True
            # Обновляем city_price для трала
            if city_name:
                from app.services.datasets import load_trail_cities
                trail_cities = load_trail_cities()
                for city in trail_cities:
                    if city.get('name') == city_name:
                        trail_price = city.get('trail_price')
                        if trail_price and trail_price > 0:
                            city_price = trail_price
                        break
        elif city_name and city_name not in TRAIL_AVAILABLE_CITIES:
            return {
                'error': f'Груз превышает грузоподъемность фуры (20 тонн), но трал доступен только до городов: {", ".join(TRAIL_AVAILABLE_CITIES)}. Для города {city_name} трал недоступен.',
                'trail_not_available': True,
                'heavy': True,
                'weight_kg': weight_kg,
            }
    
    # Определяем базис расчета (вес или объем)
    basis = determine_calculation_basis(weight_kg, volume_m3, truck_capacity, truck_volume)
    
    if basis == 'weight':
        result = calculate_logistics_by_weight(weight_kg, city_price, truck_capacity)
    else:
        result = calculate_logistics_by_volume(volume_m3, city_price, truck_volume)
    
    # Добавляем информацию о габаритах
    result.update({
        'calculation_type': 'direct',
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
        'auto_switched_to_trail': auto_switched_to_trail,
        'distance_from_ekb_provided': False,
        'can_save_distance': distance_from_ekb_km is None and not is_main_route,
    })
    
    return result


