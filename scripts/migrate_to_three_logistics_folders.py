"""
Скрипт миграции данных логистики из одного файла в 3 отдельных справочника.

Разделяет города на:
1. Основные города и города в радиусе 300км (logistics_main_cities.json)
2. Города за 300км - алгоритм ЕКБ+РФ (logistics_ekb_rf_cities.json)
3. Справочник трала (logistics_trail_cities.json)

Строгое разделение - каждый город попадает только в один справочник.
Приоритет: Трал > Основные > ЕКБ+РФ
"""

import json
import sys
from pathlib import Path

# Добавляем корневую директорию проекта в путь
BASE_DIR = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(BASE_DIR))

CONFIG_DIR = BASE_DIR / 'config'
OLD_LOGISTICS_FILE = CONFIG_DIR / 'logistics_cities.json'
MAIN_CITIES_FILE = CONFIG_DIR / 'logistics_main_cities.json'
EKB_RF_CITIES_FILE = CONFIG_DIR / 'logistics_ekb_rf_cities.json'
TRAIL_CITIES_FILE = CONFIG_DIR / 'logistics_trail_cities.json'


def migrate_logistics_cities():
    """Мигрирует города из одного файла в 3 справочника."""
    
    # Загружаем старые данные
    if not OLD_LOGISTICS_FILE.exists():
        print(f"Ошибка: Файл {OLD_LOGISTICS_FILE} не найден!", file=sys.stderr)
        return False
    
    with OLD_LOGISTICS_FILE.open('r', encoding='utf-8') as f:
        old_data = json.load(f)
    
    cities = old_data.get('cities', [])
    
    # Инициализируем списки для новых справочников
    main_cities = []
    ekb_rf_cities = []
    trail_cities = []
    
    # Статистика
    stats = {
        'total': len(cities),
        'main': 0,
        'ekb_rf': 0,
        'trail': 0,
        'skipped': 0
    }
    
    for city in cities:
        name = city.get('name', '').strip()
        if not name:
            stats['skipped'] += 1
            continue
        
        # Строгое разделение с приоритетами:
        # 1. Если allows_trail = True -> идет в справочник трала
        # 2. Если is_main_route = True или есть main_city -> идет в основные
        # 3. Иначе -> идет в ЕКБ+РФ
        
        allows_trail = city.get('allows_trail', False)
        is_main_route = city.get('is_main_route', False)
        main_city = city.get('main_city')
        
        if allows_trail:
            # Город для трала
            trail_cities.append({
                'name': name,
                'region': city.get('region', ''),
                'trail_price': city.get('trail_price') if city.get('trail_price') is not None else 0
            })
            stats['trail'] += 1
        elif is_main_route or main_city:
            # Основной город или город в радиусе 300км
            main_city_data = {
                'name': name,
                'region': city.get('region', ''),
                'truck_price': city.get('truck_price', 0),
                'is_main_route': is_main_route,
                'main_city': main_city if main_city else None
            }
            # Убираем None значения для main_city если его нет
            if main_city_data['main_city'] is None:
                del main_city_data['main_city']
            main_cities.append(main_city_data)
            stats['main'] += 1
        else:
            # Город за 300км (ЕКБ+РФ)
            ekb_rf_city_data = {
                'name': name,
                'region': city.get('region', ''),
                'distance_from_ekb_km': city.get('distance_from_ekb_km') if city.get('distance_from_ekb_km') is not None else None
            }
            # Убираем None значения для distance_from_ekb_km если его нет
            if ekb_rf_city_data['distance_from_ekb_km'] is None:
                del ekb_rf_city_data['distance_from_ekb_km']
            ekb_rf_cities.append(ekb_rf_city_data)
            stats['ekb_rf'] += 1
    
    # Сохраняем новые файлы
    CONFIG_DIR.mkdir(parents=True, exist_ok=True)
    
    # Основные города
    main_data = {'cities': main_cities}
    with MAIN_CITIES_FILE.open('w', encoding='utf-8') as f:
        json.dump(main_data, f, ensure_ascii=False, indent=2)
    print(f"[OK] Создан файл {MAIN_CITIES_FILE.name}: {len(main_cities)} городов")
    
    # ЕКБ+РФ города
    ekb_rf_data = {'cities': ekb_rf_cities}
    with EKB_RF_CITIES_FILE.open('w', encoding='utf-8') as f:
        json.dump(ekb_rf_data, f, ensure_ascii=False, indent=2)
    print(f"[OK] Создан файл {EKB_RF_CITIES_FILE.name}: {len(ekb_rf_cities)} городов")
    
    # Трал города
    trail_data = {'cities': trail_cities}
    with TRAIL_CITIES_FILE.open('w', encoding='utf-8') as f:
        json.dump(trail_data, f, ensure_ascii=False, indent=2)
    print(f"[OK] Создан файл {TRAIL_CITIES_FILE.name}: {len(trail_cities)} городов")
    
    # Выводим статистику
    print("\nСтатистика миграции:")
    print(f"  Всего городов: {stats['total']}")
    print(f"  Основные города и 300км радиус: {stats['main']}")
    print(f"  Города ЕКБ+РФ: {stats['ekb_rf']}")
    print(f"  Города трала: {stats['trail']}")
    print(f"  Пропущено (без названия): {stats['skipped']}")
    print(f"  Проверка: {stats['main'] + stats['ekb_rf'] + stats['trail'] + stats['skipped']} = {stats['total']}")
    
    if stats['main'] + stats['ekb_rf'] + stats['trail'] + stats['skipped'] != stats['total']:
        print("\n[WARNING] ВНИМАНИЕ: Сумма не совпадает с общим количеством!", file=sys.stderr)
        return False
    
    print("\n[OK] Миграция успешно завершена!")
    print("\n[WARNING] РЕКОМЕНДАЦИЯ: Создайте резервную копию старого файла перед удалением:")
    print(f"  cp {OLD_LOGISTICS_FILE} {OLD_LOGISTICS_FILE}.backup")
    
    return True


if __name__ == '__main__':
    try:
        success = migrate_logistics_cities()
        sys.exit(0 if success else 1)
    except Exception as e:
        print(f"\nОшибка при миграции: {e}", file=sys.stderr)
        import traceback
        traceback.print_exc()
        sys.exit(1)

