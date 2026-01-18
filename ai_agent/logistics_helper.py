"""
Модуль-обертка для расчета логистики.
Упрощает использование calculate_logistics для AI агента.
"""

import sys
from pathlib import Path
from typing import Any, Dict, Optional

# Добавляем корень проекта в путь для импорта
BASE_DIR = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(BASE_DIR))

from app.services import datasets
from app.services.logistics_calculator import calculate_logistics


class LogisticsHelper:
    """Класс-помощник для расчета логистики."""
    
    def __init__(self):
        """Инициализирует помощник логистики."""
        self.cities_cache: Optional[list] = None
    
    def _load_cities(self) -> list:
        """Загружает список городов (с кэшированием)."""
        if self.cities_cache is None:
            try:
                self.cities_cache = datasets.load_all_logistics_cities()
            except Exception as e:
                print(f"Ошибка при загрузке городов: {e}")
                self.cities_cache = []
        return self.cities_cache
    
    def find_city(self, city_name: str) -> Optional[Dict[str, Any]]:
        """
        Находит город по названию (улучшенный нечеткий поиск).
        
        Args:
            city_name: Название города для поиска
            
        Returns:
            Словарь с данными города или None
        """
        import logging
        logger = logging.getLogger(__name__)
        
        cities = self._load_cities()
        city_name_lower = city_name.lower().strip()
        
        if not city_name_lower:
            logger.warning("Пустое название города в find_city")
            return None
        
        # Шаг 1: Точное совпадение (приоритет)
        for city in cities:
            city_name_db = city.get('name', '').lower().strip()
            if city_name_db == city_name_lower:
                logger.debug(f"Точное совпадение: '{city_name}' -> '{city.get('name')}'")
                return city
        
        # Шаг 2: Проверка похожих названий (избегаем путаницы)
        # Словарь похожих городов, которые часто путают
        similar_cities_map = {
            'пермь': ['Пермь'],
            'чита': ['Чита'],
            'читу': ['Чита'],
            'читай': ['Чита'],
        }
        
        # Если запрос может быть перепутан, проверяем вручную
        if city_name_lower in similar_cities_map:
            expected_cities = similar_cities_map[city_name_lower]
            for city in cities:
                if city.get('name') in expected_cities:
                    logger.debug(f"Найден город по словарю похожих: '{city_name}' -> '{city.get('name')}'")
                    return city
        
        # Шаг 3: Начинается с (более точное совпадение)
        matches_start_with = []
        for city in cities:
            city_name_db = city.get('name', '').lower().strip()
            if city_name_db.startswith(city_name_lower) or city_name_lower.startswith(city_name_db):
                matches_start_with.append((len(city_name_db), city))
        
        if matches_start_with:
            # Возвращаем самый длинный совпадение (самое полное название)
            matches_start_with.sort(key=lambda x: x[0], reverse=True)
            best_match = matches_start_with[0][1]
            logger.debug(f"Найден город по началу: '{city_name}' -> '{best_match.get('name')}'")
            return best_match
        
        # Шаг 4: Частичное совпадение (fallback)
        matches_partial = []
        for city in cities:
            city_name_db = city.get('name', '').lower().strip()
            if city_name_lower in city_name_db or city_name_db in city_name_lower:
                # Приоритет более длинным названиям
                matches_partial.append((len(city_name_db), city))
        
        if matches_partial:
            # Возвращаем самое длинное совпадение
            matches_partial.sort(key=lambda x: x[0], reverse=True)
            best_match = matches_partial[0][1]
            logger.debug(f"Найден город по частичному совпадению: '{city_name}' -> '{best_match.get('name')}'")
            return best_match
        
        logger.warning(f"Город '{city_name}' не найден в справочнике")
        return None
    
    def calculate_simple_logistics(
        self,
        weight_kg: float,
        city_name: str,
        transport_type: str = 'truck',
        length_mm: Optional[float] = None,
        width_mm: Optional[float] = None,
        height_mm: Optional[float] = None,
    ) -> Dict[str, Any]:
        """
        Упрощенный расчет логистики.
        
        Args:
            weight_kg: Вес груза в кг
            city_name: Название города назначения
            transport_type: Тип транспорта ('truck' или 'trail')
            length_mm: Длина груза в мм (опционально)
            width_mm: Ширина груза в мм (опционально)
            height_mm: Высота груза в мм (опционально)
            
        Returns:
            Словарь с результатами расчета или ошибкой
        """
        # Находим город
        city = self.find_city(city_name)
        if not city:
            return {
                'error': f'Город "{city_name}" не найден в справочнике. Проверьте название города.',
                'available_cities': [c.get('name') for c in self._load_cities()[:10]]
            }
        
        # Извлекаем параметры города
        city_price = city.get('truck_price', 0)
        is_main_route = city.get('is_main_route', False)
        distance_from_ekb_km = city.get('distance_from_ekb_km')
        allows_trail = city.get('allows_trail', False)
        city_in_ekb_rf = datasets.is_city_in_ekb_rf_catalog(city.get('name', ''))
        
        # Если трал запрошен, но недоступен
        if transport_type == 'trail' and not allows_trail:
            return {
                'error': f'Трал недоступен для города "{city_name}". Используйте тип транспорта "truck".',
                'city_name': city.get('name')
            }
        
        # Если трал запрошен, используем trail_price
        if transport_type == 'trail' and allows_trail:
            city_price = city.get('trail_price', city_price)
        
        # Если расстояние не указано, но город в ЕКБ+РФ, получаем из справочника
        if distance_from_ekb_km is None and city_in_ekb_rf:
            distance_from_ekb_km = datasets.get_ekb_rf_city_distance(city.get('name', ''))
        
        # Выполняем расчет
        try:
            result = calculate_logistics(
                weight_kg=weight_kg,
                city_price=city_price,
                transport_type=transport_type,
                city_name=city.get('name'),
                is_main_route=is_main_route,
                distance_from_ekb_km=distance_from_ekb_km,
                length_mm=length_mm,
                width_mm=width_mm,
                height_mm=height_mm,
                city_in_ekb_rf_catalog=city_in_ekb_rf if city_in_ekb_rf else None,
            )
            
            # Добавляем информацию о городе
            result['city_name'] = city.get('name')
            result['region'] = city.get('region')
            
            return result
        except Exception as e:
            return {
                'error': f'Ошибка при расчете логистики: {str(e)}',
                'city_name': city.get('name')
            }
    
    def format_logistics_result(self, result: Dict[str, Any]) -> str:
        """
        Форматирует результат расчета логистики для вывода.
        Соответствует формату страницы логистики.
        
        Args:
            result: Результат расчета из calculate_logistics
            
        Returns:
            Отформатированная строка с результатами
        """
        def format_number(num: float) -> str:
            """Форматирует число с пробелами в качестве разделителей тысяч."""
            return f"{num:,.0f}".replace(",", " ")
        
        def format_money(num: float) -> str:
            """Форматирует денежную сумму с пробелами в качестве разделителей тысяч."""
            return format_number(num)
        
        if 'error' in result:
            error_msg = f"❌ Ошибка: {result['error']}"
            if 'available_cities' in result:
                cities = ', '.join(result['available_cities'])
                error_msg += f"\n\nДоступные города (первые 10): {cities}"
            return error_msg
        
        if result.get('basis') == 'small_cargo':
            weight_kg = result.get('weight_kg', 0)
            return (
                f"📦 Мелкогабаритный груз (≤ 600 кг)\n\n"
                f"Вес груза: {format_number(weight_kg)} кг\n\n"
                f"Рекомендация: {result.get('recommendation', 'dellin')}\n"
                f"{result.get('message', '')}"
            )
        
        # Основная информация
        city_name = result.get('city_name', 'Неизвестно')
        region = result.get('region', '')
        total_price = result.get('total_price', 0)
        basis = result.get('basis', 'weight')
        calculation_type = result.get('calculation_type', 'direct')
        
        output = [
            "🚚 Расчет логистики",
            f"Город: {city_name}" + (f" ({region})" if region else ""),
            f"Вес: {format_number(result.get('weight_kg', 0))} кг",
        ]
        
        # Добавляем объем, если есть
        if 'volume_m3' in result:
            output.append(f"Объем: {result['volume_m3']:.2f} м³")
        
        output.append("")  # Пустая строка перед алгоритмом
        
        # Тип расчета и направление
        if calculation_type == 'ekb_plus_rf':
            output.append("📊 Алгоритм: КНР → Екатеринбург → Город")
            route_details = result.get('route_details', {})
            if route_details:
                china_to_ekb = route_details.get('china_to_ekb', {})
                ekb_to_dest = route_details.get('ekb_to_destination', {})
                output.append(f"  КНР → ЕКБ: {format_money(china_to_ekb.get('price', 0))} руб")
                output.append(f"  ЕКБ → {city_name}: {format_money(ekb_to_dest.get('price', 0))} руб")
        else:
            output.append("📊 Алгоритм: Прямой расчет")
            output.append(f"Направление: КНР → {city_name}")
        
        # Базис расчета
        if basis == 'weight':
            price_per_kg = result.get('price_per_kg', 0)
            output.append("Базис: по весу")
            output.append(f"Цена за кг: {price_per_kg:,.2f} руб/кг")
        elif basis == 'volume':
            price_per_m3 = result.get('price_per_m3', 0)
            output.append("Базис: по объему")
            output.append(f"Цена за м³: {price_per_m3:,.2f} руб/м³")
        
        # Количество машин
        trucks_count = result.get('trucks_count', 1)
        output.append(f"Количество машин: {trucks_count:.1f}")
        
        # Итоговая стоимость
        output.append("")
        output.append(f"💰 ИТОГО: {format_money(total_price)} руб")
        
        # Формула расчета
        if 'calculation_formula' in result:
            output.append("")
            output.append(f"📐 Формула: {result['calculation_formula']}")
        
        return "\n".join(output)

