"""
Скрипт миграции данных logistics_cities.json для добавления новых полей.

Добавляет:
- is_main_route: bool (город на основном маршруте КНР → Россия)
- distance_from_ekb_km: int | null (расстояние от Екатеринбурга)
- allows_trail: bool (доступен ли трал для города)

Основные города маршрута согласно документу:
Чита, Улан-Удэ, Иркутск, Красноярск, Новосибирск, Омск, Екатеринбург, Москва, Санкт-Петербург, Минск
"""

import json
from pathlib import Path

# Основные города маршрута КНР → Россия → Беларусь
MAIN_ROUTE_CITIES = [
    'Чита', 'Улан-Удэ', 'Иркутск', 'Красноярск', 'Новосибирск', 
    'Омск', 'Екатеринбург', 'Москва', 'Санкт-Петербург', 'Минск'
]

# Города, где доступен трал (только Екатеринбург и Москва)
TRAIL_CITIES = ['Екатеринбург', 'Москва']

# Тарифы для основных городов (из документа раздел 3.2 и 3.3)
MAIN_CITY_PRICES = {
    'Чита': {'truck': 400000, 'trail': None},
    'Улан-Удэ': {'truck': 500000, 'trail': None},
    'Иркутск': {'truck': 600000, 'trail': None},
    'Красноярск': {'truck': 800000, 'trail': None},
    'Новосибирск': {'truck': 900000, 'trail': None},
    'Омск': {'truck': 1000000, 'trail': None},
    'Екатеринбург': {'truck': 1100000, 'trail': 2000000},
    'Москва': {'truck': 1200000, 'trail': 2100000},
    'Санкт-Петербург': {'truck': 1300000, 'trail': None},
    'Минск': {'truck': 1300000, 'trail': None},
}


def migrate_city(city: dict) -> dict:
    """Обновляет структуру одного города."""
    city_name = city.get('name', '')
    
    # Определяем, является ли город основным на маршруте
    is_main_route = city_name in MAIN_ROUTE_CITIES
    
    # Определяем, доступен ли трал
    allows_trail = city_name in TRAIL_CITIES
    
    # Обновляем цены для основных городов
    if city_name in MAIN_CITY_PRICES:
        prices = MAIN_CITY_PRICES[city_name]
        city['truck_price'] = prices['truck']
        if prices['trail'] is not None:
            city['trail_price'] = prices['trail']
        else:
            city['trail_price'] = None
    
    # Для не основных городов убираем trail_price если трал недоступен
    if not allows_trail:
        city['trail_price'] = None
    
    # Добавляем новые поля
    city['is_main_route'] = is_main_route
    city['allows_trail'] = allows_trail
    city['distance_from_ekb_km'] = None  # Будет заполняться пользователями
    
    # Удаляем устаревшее поле is_main (заменено на is_main_route)
    if 'is_main' in city:
        del city['is_main']
    
    return city


def main():
    """Основная функция миграции."""
    config_path = Path(__file__).resolve().parents[1] / 'config' / 'logistics_cities.json'
    
    print(f'Читаем файл: {config_path}')
    
    with config_path.open('r', encoding='utf-8') as f:
        data = json.load(f)
    
    cities = data.get('cities', [])
    print(f'Найдено городов: {len(cities)}')
    
    # Мигрируем все города
    migrated_cities = []
    for city in cities:
        migrated_city = migrate_city(city)
        migrated_cities.append(migrated_city)
    
    # Сохраняем обновленные данные
    data['cities'] = migrated_cities
    
    # Создаем backup
    backup_path = config_path.with_suffix('.json.backup')
    print(f'Создаем backup: {backup_path}')
    with backup_path.open('w', encoding='utf-8') as f:
        json.dump(data, f, ensure_ascii=False, indent=2)
    
    # Сохраняем обновленный файл
    print(f'Сохраняем обновленный файл: {config_path}')
    with config_path.open('w', encoding='utf-8') as f:
        json.dump(data, f, ensure_ascii=False, indent=2)
    
    # Статистика
    main_route_count = sum(1 for c in migrated_cities if c.get('is_main_route'))
    trail_cities_count = sum(1 for c in migrated_cities if c.get('allows_trail'))
    
    print('\nМиграция завершена успешно!')
    print(f'  - Всего городов: {len(migrated_cities)}')
    print(f'  - Городов на основном маршруте: {main_route_count}')
    print(f'  - Городов с доступным тралом: {trail_cities_count}')


if __name__ == '__main__':
    main()

