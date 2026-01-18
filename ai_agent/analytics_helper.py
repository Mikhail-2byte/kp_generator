"""
Модуль для работы с данными из 1С (CSV файлы).
Предоставляет поиск по номенклатуре, чертежам, заказчикам и генерацию графиков.
"""

import base64
import csv
import logging
import re
from io import BytesIO
from pathlib import Path
from typing import Any, Dict, List, Optional

import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt
import numpy as np
import pandas as pd
import seaborn as sns

from ai_agent.formatters import ResponseFormatter
from ai_agent.suggestions import SuggestionsGenerator

sns.set_theme(style='whitegrid')

logger = logging.getLogger(__name__)


class AnalyticsHelper:
    """Класс для работы с данными из 1С (CSV файлы)."""
    
    def __init__(self, lazy_load: bool = True):
        """
        Инициализирует помощник аналитики.
        
        Args:
            lazy_load: Если True, данные загружаются при первом запросе, иначе при инициализации
        """
        self.data_dir = Path(__file__).resolve().parent / "data" / "1с"
        self.data_cache: Dict[str, pd.DataFrame] = {}
        self.categories: List[str] = []
        self._data_loaded = False
        self._lazy_load = lazy_load
        
        # Загружаем данные только если не включена ленивая загрузка
        if not lazy_load:
            self._load_all_data()
    
    def _load_all_data(self) -> None:
        """
        Загружает все CSV файлы из папки 1с/.
        Каждый файл становится категорией (например, Барабаны, Вал).
        При загрузке данных инвалидирует кеш аналитики.
        """
        if not self.data_dir.exists():
            logger.warning(f"Папка с данными 1С не найдена: {self.data_dir}")
            return
        
        # Ищем все CSV файлы
        csv_files = list(self.data_dir.glob("*.csv"))
        
        if not csv_files:
            logger.warning(f"CSV файлы в папке {self.data_dir} не найдены")
            return
        
        # Инвалидируем кеш аналитики при загрузке новых данных
        try:
            from ai_agent.cache_manager import invalidate_analytics_cache
            invalidate_analytics_cache()
            logger.info("Кеш аналитики инвалидирован при загрузке данных")
        except Exception as e:
            logger.warning(f"Не удалось инвалидировать кеш аналитики: {e}")
        
        for csv_file in csv_files:
            category = csv_file.stem  # Имя файла без расширения (например, "Барабаны")
            try:
                df = self._load_csv(csv_file)
                if df is not None and not df.empty:
                    self.data_cache[category] = df
                    self.categories.append(category)
                    logger.info(f"Загружено {len(df)} записей из категории '{category}'")
            except Exception as e:
                logger.error(f"Ошибка при загрузке файла {csv_file}: {e}", exc_info=True)
    
    def _load_csv(self, csv_path: Path) -> Optional[pd.DataFrame]:
        """
        Загружает и обрабатывает CSV файл из 1С с улучшенной обработкой ошибок.
        
        Args:
            csv_path: Путь к CSV файлу
            
        Returns:
            DataFrame с обработанными данными или None при ошибке
        """
        try:
            # Пробуем разные кодировки
            encodings = ['utf-8-sig', 'utf-8', 'cp1251', 'windows-1251']
            df = None
            encoding_used = None
            
            for encoding in encodings:
                try:
                    df = pd.read_csv(
                        csv_path,
                        encoding=encoding,
                        quotechar='"',
                        delimiter=',',
                        skipinitialspace=True,
                        on_bad_lines='skip'  # Пропускаем проблемные строки
                    )
                    encoding_used = encoding
                    break
                except UnicodeDecodeError:
                    continue
                except Exception as e:
                    logger.warning(f"Ошибка при чтении CSV {csv_path} с кодировкой {encoding}: {e}")
                    continue
            
            if df is None or df.empty:
                logger.error(f"Не удалось загрузить CSV файл {csv_path}: файл пуст или нечитаем")
                return None
            
            logger.debug(f"CSV {csv_path} загружен с кодировкой {encoding_used}, строк: {len(df)}")
            
            # Валидация структуры CSV
            validation_result = self._validate_csv_structure(df, csv_path)
            if not validation_result['is_valid']:
                logger.error(f"CSV {csv_path} не прошел валидацию: {validation_result['errors']}")
                # Пытаемся автоматически исправить распространенные ошибки
                if validation_result.get('can_fix', False):
                    df = self._fix_csv_structure(df, validation_result.get('fixes', []))
                    logger.info(f"Применены автоматические исправления к CSV {csv_path}")
                else:
                    return None
            
            # Проверяем наличие обязательных колонок
            required_columns = ['Чертеж']
            missing_columns = [col for col in required_columns if col not in df.columns]
            if missing_columns:
                logger.error(f"В CSV {csv_path} отсутствуют обязательные колонки: {missing_columns}")
                logger.info(f"Доступные колонки: {list(df.columns)}")
                return None
            
            # Пропускаем строки без чертежа (групповые строки)
            initial_count = len(df)
            df = df[df['Чертеж'].notna() & (df['Чертеж'].astype(str).str.strip() != '')]
            filtered_count = len(df)
            
            if filtered_count < initial_count:
                logger.debug(f"Отфильтровано {initial_count - filtered_count} строк без чертежа")
            
            if df.empty:
                logger.warning(f"После фильтрации CSV {csv_path} не содержит записей")
                return None
            
            # Нормализуем числовые колонки с обработкой ошибок
            numeric_columns = [
                'Количество', 'ГП вес, кг', 'Цена подачи', 
                'Цена подачи за кг', 'Мин цена подачи за кг',
                'Итоговый вес', 'Минимальная цена закупа за кг.'
            ]
            
            for col in numeric_columns:
                if col in df.columns:
                    try:
                        df[col] = self._normalize_numeric(df[col])
                        # Проверяем, что нормализация прошла успешно
                        if df[col].isna().all():
                            logger.warning(f"Колонка {col} не содержит валидных числовых значений")
                    except Exception as e:
                        logger.warning(f"Ошибка при нормализации колонки {col} в {csv_path}: {e}")
                        # Продолжаем работу, даже если одна колонка не нормализовалась
            
            # Парсим дату и статус из колонки "Тендер Дата выхода тендера, Статус"
            if 'Тендер Дата выхода тендера, Статус' in df.columns:
                try:
                    df['Дата_тендера'] = df['Тендер Дата выхода тендера, Статус'].astype(str).str.split(',').str[0].str.strip()
                    df['Статус_тендера'] = df['Тендер Дата выхода тендера, Статус'].astype(str).str.split(',').str[1:].str.join(',').str.strip()
                    
                    # Парсим дату с обработкой ошибок
                    df['Дата_тендера_parsed'] = pd.to_datetime(df['Дата_тендера'], errors='coerce', dayfirst=True)
                    
                    # Логируем количество успешно распарсенных дат
                    parsed_dates_count = df['Дата_тендера_parsed'].notna().sum()
                    if parsed_dates_count < len(df):
                        logger.debug(f"Успешно распарсено {parsed_dates_count} из {len(df)} дат")
                except Exception as e:
                    logger.warning(f"Ошибка при парсинге дат тендеров в {csv_path}: {e}")
                    # Продолжаем работу без дат
            
            # Нормализуем названия колонок для удобства
            df.columns = df.columns.str.strip()
            
            logger.info(f"CSV {csv_path} успешно обработан: {len(df)} записей")
            return df
            
        except pd.errors.EmptyDataError:
            logger.error(f"CSV файл {csv_path} пуст")
            return None
        except pd.errors.ParserError as e:
            logger.error(f"Ошибка парсинга CSV {csv_path}: {e}")
            return None
        except FileNotFoundError:
            logger.error(f"CSV файл {csv_path} не найден")
            return None
        except PermissionError:
            logger.error(f"Нет доступа к CSV файлу {csv_path}")
            return None
        except Exception as e:
            logger.error(f"Неожиданная ошибка при парсинге CSV {csv_path}: {e}", exc_info=True)
            return None
    
    def _validate_csv_structure(self, df: pd.DataFrame, csv_path: Path) -> Dict[str, Any]:
        """
        Валидирует структуру CSV файла.
        
        Args:
            df: DataFrame для валидации
            csv_path: Путь к CSV файлу (для логирования)
            
        Returns:
            Словарь с результатами валидации
        """
        errors = []
        warnings = []
        fixes = []
        can_fix = False
        
        # Проверка на пустой DataFrame
        if df.empty:
            errors.append("CSV файл пуст")
            return {
                'is_valid': False,
                'errors': errors,
                'warnings': warnings,
                'can_fix': False,
                'fixes': fixes
            }
        
        # Проверка наличия колонок
        expected_columns = [
            'Номенклатура', 'Чертеж', 'Заказчик, сокр', 'Количество',
            'ГП вес, кг', 'Материал', 'Цена подачи', 'Менеджер',
            'Тендер Дата выхода тендера, Статус'
        ]
        
        missing_columns = [col for col in expected_columns if col not in df.columns]
        if missing_columns:
            # Проверяем, есть ли похожие колонки (с пробелами, без пробелов)
            for missing_col in missing_columns:
                # Ищем похожие колонки
                similar_cols = [col for col in df.columns if missing_col.lower().replace(' ', '') in col.lower().replace(' ', '')]
                if similar_cols:
                    warnings.append(f"Колонка '{missing_col}' не найдена, но найдены похожие: {similar_cols}")
                    can_fix = True
                    fixes.append({
                        'type': 'column_rename',
                        'from': similar_cols[0],
                        'to': missing_col
                    })
                else:
                    if missing_col == 'Чертеж':
                        errors.append(f"Обязательная колонка '{missing_col}' отсутствует")
                    else:
                        warnings.append(f"Колонка '{missing_col}' отсутствует (необязательная)")
        
        # Проверка типов данных
        if 'Количество' in df.columns:
            # Проверяем, что можно преобразовать в число
            try:
                sample = df['Количество'].head(10)
                pd.to_numeric(sample, errors='raise')
            except (ValueError, TypeError):
                warnings.append("Колонка 'Количество' содержит нечисловые значения")
                can_fix = True
                fixes.append({'type': 'normalize_numeric', 'column': 'Количество'})
        
        # Проверка на дубликаты колонок
        if len(df.columns) != len(set(df.columns)):
            duplicates = [col for col in df.columns if df.columns.tolist().count(col) > 1]
            warnings.append(f"Обнаружены дубликаты колонок: {duplicates}")
            can_fix = True
            fixes.append({'type': 'fix_duplicate_columns'})
        
        is_valid = len(errors) == 0
        
        return {
            'is_valid': is_valid,
            'errors': errors,
            'warnings': warnings,
            'can_fix': can_fix,
            'fixes': fixes
        }
    
    def _fix_csv_structure(self, df: pd.DataFrame, fixes: List[Dict[str, Any]]) -> pd.DataFrame:
        """
        Применяет автоматические исправления к структуре CSV.
        
        Args:
            df: DataFrame для исправления
            fixes: Список исправлений
            
        Returns:
            Исправленный DataFrame
        """
        df_fixed = df.copy()
        
        for fix in fixes:
            try:
                if fix['type'] == 'column_rename':
                    if fix['from'] in df_fixed.columns:
                        df_fixed = df_fixed.rename(columns={fix['from']: fix['to']})
                        logger.debug(f"Переименована колонка '{fix['from']}' в '{fix['to']}'")
                
                elif fix['type'] == 'normalize_numeric':
                    col = fix['column']
                    if col in df_fixed.columns:
                        df_fixed[col] = self._normalize_numeric(df_fixed[col])
                        logger.debug(f"Нормализована числовая колонка '{col}'")
                
                elif fix['type'] == 'fix_duplicate_columns':
                    # Удаляем дубликаты колонок, оставляя первую
                    seen = set()
                    cols_to_keep = []
                    for col in df_fixed.columns:
                        if col not in seen:
                            seen.add(col)
                            cols_to_keep.append(col)
                    df_fixed = df_fixed[cols_to_keep]
                    logger.debug("Удалены дубликаты колонок")
            except Exception as e:
                logger.warning(f"Не удалось применить исправление {fix}: {e}")
                continue
        
        return df_fixed
    
    def _normalize_numeric(self, series: pd.Series) -> pd.Series:
        """
        Нормализует числовую серию, заменяя запятые на точки и убирая пробелы.
        
        Args:
            series: Серия pandas
            
        Returns:
            Нормализованная серия с числовыми значениями
        """
        if series.dtype == object:
            cleaned = (
                series.astype(str)
                .str.replace('\xa0', '', regex=False)  # Неразрывные пробелы
                .str.replace(' ', '', regex=False)  # Обычные пробелы
                .str.replace(',', '.', regex=False)  # Запятые на точки
            )
            return pd.to_numeric(cleaned, errors='coerce')
        return pd.to_numeric(series, errors='coerce')
    
    def _ensure_data_loaded(self) -> None:
        """Обеспечивает загрузку данных, если они еще не загружены."""
        if not self._data_loaded:
            logger.info("Загрузка данных 1С при первом запросе...")
            self._load_all_data()
            self._data_loaded = True
    
    def get_categories(self) -> List[str]:
        """Возвращает список доступных категорий."""
        self._ensure_data_loaded()
        return self.categories.copy()
    
    def search_by_nomenclature(
        self, 
        query: str, 
        category: Optional[str] = None, 
        limit: Optional[int] = None,
        filters: Optional[Dict[str, Any]] = None,
        sort_by: Optional[str] = None,
        sort_order: str = 'asc'
    ) -> pd.DataFrame:
        """
        Поиск по номенклатуре с обработкой ошибок.
        
        Args:
            query: Поисковый запрос
            category: Категория для поиска (если None, ищет во всех)
            limit: Максимальное количество результатов (если None, возвращает все)
            
        Returns:
            DataFrame с найденными записями
        """
        self._ensure_data_loaded()
        
        if not query or not query.strip():
            logger.warning("Пустой запрос для поиска по номенклатуре")
            return pd.DataFrame()
        
        query_lower = query.lower().strip()
        results = []
        errors = []
        
        categories_to_search = [category] if category and category in self.data_cache else self.categories
        
        if not categories_to_search:
            logger.warning("Нет доступных категорий для поиска")
            return pd.DataFrame()
        
        for cat in categories_to_search:
            try:
                if cat not in self.data_cache:
                    logger.warning(f"Категория {cat} не найдена в кеше")
                    continue
                
                df = self.data_cache[cat]
                if df.empty:
                    continue
                
                if 'Номенклатура' not in df.columns:
                    logger.warning(f"В категории {cat} отсутствует колонка 'Номенклатура'")
                    continue
                
                try:
                    mask = df['Номенклатура'].astype(str).str.lower().str.contains(query_lower, na=False)
                    found = df[mask].copy()
                    if not found.empty:
                        found['Категория'] = cat
                        results.append(found)
                except Exception as e:
                    logger.warning(f"Ошибка при поиске в категории {cat}: {e}")
                    errors.append(f"Ошибка в категории {cat}")
                    continue
            except Exception as e:
                logger.error(f"Критическая ошибка при обработке категории {cat}: {e}", exc_info=True)
                errors.append(f"Критическая ошибка в категории {cat}")
                continue
        
        if results:
            try:
                result_df = pd.concat(results, ignore_index=True)
                if errors:
                    logger.info(f"Поиск завершен с предупреждениями: {', '.join(errors)}")
                
                # Применяем фильтры
                if filters:
                    result_df = self.filter_results(result_df, filters)
                
                # Применяем сортировку
                if sort_by:
                    result_df = self.sort_results(result_df, sort_by, sort_order)
                
                # Применяем лимит, если указан (после фильтрации и сортировки)
                if limit is not None and limit > 0:
                    result_df = result_df.head(limit)
                return result_df
            except Exception as e:
                logger.error(f"Ошибка при объединении результатов поиска: {e}", exc_info=True)
                # Возвращаем первый результат, если объединение не удалось
                if results:
                    return results[0]
                return pd.DataFrame()
        
        if errors:
            logger.warning(f"Поиск не дал результатов из-за ошибок: {', '.join(errors)}")
        
        return pd.DataFrame()
    
    def search_by_drawing(
        self, 
        drawing: str, 
        category: Optional[str] = None, 
        limit: Optional[int] = None,
        filters: Optional[Dict[str, Any]] = None,
        sort_by: Optional[str] = None,
        sort_order: str = 'asc'
    ) -> pd.DataFrame:
        """
        Поиск по чертежу (точный и частичный поиск).
        
        Args:
            drawing: Номер чертежа
            category: Категория для поиска (если None, ищет во всех)
            limit: Максимальное количество результатов (если None, возвращает все)
            
        Returns:
            DataFrame с найденными записями
        """
        self._ensure_data_loaded()
        drawing_clean = drawing.strip()
        results = []
        
        categories_to_search = [category] if category and category in self.data_cache else self.categories
        
        for cat in categories_to_search:
            df = self.data_cache[cat]
            if 'Чертеж' in df.columns:
                # Точное совпадение (приоритет)
                exact_mask = df['Чертеж'].astype(str).str.strip() == drawing_clean
                if exact_mask.any():
                    found = df[exact_mask].copy()
                    found['Категория'] = cat
                    results.append(found)
                else:
                    # Частичное совпадение
                    partial_mask = df['Чертеж'].astype(str).str.contains(drawing_clean, na=False, case=False)
                    found = df[partial_mask].copy()
                    if not found.empty:
                        found['Категория'] = cat
                        results.append(found)
        
        if results:
            result_df = pd.concat(results, ignore_index=True)
            
            # Применяем фильтры
            if filters:
                result_df = self.filter_results(result_df, filters)
            
            # Применяем сортировку
            if sort_by:
                result_df = self.sort_results(result_df, sort_by, sort_order)
            
            # Применяем лимит, если указан
            if limit is not None and limit > 0:
                result_df = result_df.head(limit)
            return result_df
        return pd.DataFrame()
    
    def search_by_customer(
        self, 
        customer: str, 
        category: Optional[str] = None, 
        limit: Optional[int] = None,
        filters: Optional[Dict[str, Any]] = None,
        sort_by: Optional[str] = None,
        sort_order: str = 'asc'
    ) -> pd.DataFrame:
        """
        Поиск по заказчику.
        
        Args:
            customer: Название заказчика
            category: Категория для поиска (если None, ищет во всех)
            limit: Максимальное количество результатов (если None, возвращает все)
            
        Returns:
            DataFrame с найденными записями
        """
        self._ensure_data_loaded()
        customer_lower = customer.lower().strip()
        results = []
        
        categories_to_search = [category] if category and category in self.data_cache else self.categories
        
        for cat in categories_to_search:
            df = self.data_cache[cat]
            customer_col = 'Заказчик, сокр' if 'Заказчик, сокр' in df.columns else 'Заказчик'
            if customer_col in df.columns:
                mask = df[customer_col].astype(str).str.lower().str.contains(customer_lower, na=False)
                found = df[mask].copy()
                found['Категория'] = cat
                results.append(found)
        
        if results:
            result_df = pd.concat(results, ignore_index=True)
            
            # Применяем фильтры
            if filters:
                result_df = self.filter_results(result_df, filters)
            
            # Применяем сортировку
            if sort_by:
                result_df = self.sort_results(result_df, sort_by, sort_order)
            
            # Применяем лимит, если указан
            if limit is not None and limit > 0:
                result_df = result_df.head(limit)
            return result_df
        return pd.DataFrame()
    
    def search_by_material(
        self, 
        material: str, 
        category: Optional[str] = None, 
        limit: Optional[int] = None,
        filters: Optional[Dict[str, Any]] = None,
        sort_by: Optional[str] = None,
        sort_order: str = 'asc'
    ) -> pd.DataFrame:
        """
        Поиск по материалу.
        
        Args:
            material: Название материала
            category: Категория для поиска (если None, ищет во всех)
            limit: Максимальное количество результатов (если None, возвращает все)
            
        Returns:
            DataFrame с найденными записями
        """
        self._ensure_data_loaded()
        material_lower = material.lower().strip()
        results = []
        
        categories_to_search = [category] if category and category in self.data_cache else self.categories
        
        for cat in categories_to_search:
            df = self.data_cache[cat]
            if 'Материал' in df.columns:
                mask = df['Материал'].astype(str).str.lower().str.contains(material_lower, na=False)
                found = df[mask].copy()
                found['Категория'] = cat
                results.append(found)
        
        if results:
            result_df = pd.concat(results, ignore_index=True)
            
            # Применяем фильтры
            if filters:
                result_df = self.filter_results(result_df, filters)
            
            # Применяем сортировку
            if sort_by:
                result_df = self.sort_results(result_df, sort_by, sort_order)
            
            # Применяем лимит, если указан
            if limit is not None and limit > 0:
                result_df = result_df.head(limit)
            return result_df
        return pd.DataFrame()
    
    def search_combined(
        self,
        category: Optional[str] = None,
        nomenclature: Optional[str] = None,
        drawing: Optional[str] = None,
        customer: Optional[str] = None,
        material: Optional[str] = None,
        limit: Optional[int] = None,
        filters: Optional[Dict[str, Any]] = None,
        sort_by: Optional[str] = None,
        sort_order: str = 'asc',
        **kwargs
    ) -> pd.DataFrame:
        """
        Комбинированный поиск по нескольким параметрам с поддержкой фильтрации и сортировки.
        
        Args:
            category: Категория
            nomenclature: Номенклатура
            drawing: Чертеж
            customer: Заказчик
            material: Материал
            limit: Максимальное количество результатов (если None, возвращает все)
            filters: Дополнительные фильтры (словарь)
            sort_by: Столбец для сортировки
            sort_order: Порядок сортировки ('asc' или 'desc')
            **kwargs: Дополнительные фильтры (устаревший способ, используйте filters)
            
        Returns:
            DataFrame с найденными записями
        """
        self._ensure_data_loaded()
        categories_to_search = [category] if category and category in self.data_cache else self.categories
        
        if not categories_to_search:
            return pd.DataFrame()
        
        results = []
        
        for cat in categories_to_search:
            df = self.data_cache[cat]
            mask = pd.Series([True] * len(df), index=df.index)
            
            if nomenclature:
                mask &= df['Номенклатура'].astype(str).str.lower().str.contains(nomenclature.lower(), na=False)
            
            if drawing:
                mask &= df['Чертеж'].astype(str).str.contains(drawing, na=False, case=False)
            
            if customer:
                customer_col = 'Заказчик, сокр' if 'Заказчик, сокр' in df.columns else 'Заказчик'
                if customer_col in df.columns:
                    mask &= df[customer_col].astype(str).str.lower().str.contains(customer.lower(), na=False)
            
            if material:
                mask &= df['Материал'].astype(str).str.lower().str.contains(material.lower(), na=False)
            
            found = df[mask].copy()
            if not found.empty:
                found['Категория'] = cat
                results.append(found)
        
        if results:
            result_df = pd.concat(results, ignore_index=True)
            
            # Применяем дополнительные фильтры
            if filters:
                result_df = self.filter_results(result_df, filters)
            
            # Применяем сортировку
            if sort_by:
                result_df = self.sort_results(result_df, sort_by, sort_order)
            
            # Применяем лимит, если указан (после фильтрации и сортировки)
            if limit is not None and limit > 0:
                result_df = result_df.head(limit)
            
            return result_df
        return pd.DataFrame()
    
    def get_statistics(self, category: Optional[str] = None) -> Dict[str, Any]:
        """
        Получает общую статистику по категории или всем категориям.
        
        Args:
            category: Категория (если None, статистика по всем)
            
        Returns:
            Словарь со статистикой
        """
        self._ensure_data_loaded()
        if category and category in self.data_cache:
            df = self.data_cache[category]
        elif not self.categories:
            return {}
        else:
            # Объединяем все категории
            dfs = [self.data_cache[cat] for cat in self.categories]
            df = pd.concat(dfs, ignore_index=True)
        
        stats = {
            'total_records': len(df),
            'categories': [category] if category else self.categories
        }
        
        if 'Цена подачи' in df.columns:
            total_price = float(df['Цена подачи'].sum() if df['Цена подачи'].notna().any() else 0)
            stats['total_price'] = total_price
        
        if 'ГП вес, кг' in df.columns:
            total_weight = float(df['ГП вес, кг'].sum() if df['ГП вес, кг'].notna().any() else 0)
            stats['total_weight_kg'] = total_weight
        
        if 'Заказчик, сокр' in df.columns:
            stats['unique_customers'] = int(df['Заказчик, сокр'].nunique())
        elif 'Заказчик' in df.columns:
            stats['unique_customers'] = int(df['Заказчик'].nunique())
        
        if 'Материал' in df.columns:
            stats['unique_materials'] = int(df['Материал'].nunique())
        
        return stats
    
    def get_top_customers(self, category: Optional[str] = None, limit: int = 10) -> pd.DataFrame:
        """
        Получает топ заказчиков по количеству заказов.
        
        Args:
            category: Категория (если None, по всем)
            limit: Количество топ результатов
            
        Returns:
            DataFrame с топ заказчиками
        """
        self._ensure_data_loaded()
        if category and category in self.data_cache:
            df = self.data_cache[category]
        elif not self.categories:
            return pd.DataFrame()
        else:
            dfs = [self.data_cache[cat] for cat in self.categories]
            df = pd.concat(dfs, ignore_index=True)
        
        customer_col = 'Заказчик, сокр' if 'Заказчик, сокр' in df.columns else 'Заказчик'
        if customer_col not in df.columns:
            return pd.DataFrame()
        
        top = df[customer_col].value_counts().head(limit).reset_index()
        top.columns = ['Заказчик', 'Количество']
        
        return top
    
    def get_top_materials(self, category: Optional[str] = None, limit: int = 10) -> pd.DataFrame:
        """
        Получает топ материалов по количеству использований.
        
        Args:
            category: Категория (если None, по всем)
            limit: Количество топ результатов
            
        Returns:
            DataFrame с топ материалами
        """
        self._ensure_data_loaded()
        if category and category in self.data_cache:
            df = self.data_cache[category]
        elif not self.categories:
            return pd.DataFrame()
        else:
            dfs = [self.data_cache[cat] for cat in self.categories]
            df = pd.concat(dfs, ignore_index=True)
        
        if 'Материал' not in df.columns:
            return pd.DataFrame()
        
        top = df['Материал'].value_counts().head(limit).reset_index()
        top.columns = ['Материал', 'Количество']
        
        return top
    
    def search_by_tender_status(self, status: str, category: Optional[str] = None, limit: Optional[int] = None) -> pd.DataFrame:
        """
        Поиск по статусу тендера.
        
        Args:
            status: Статус тендера (Выиграли, Проиграли, Подались, Отменен, Выиграли частично, Не участвовали)
            category: Категория (если None, по всем)
            limit: Максимальное количество результатов (если None, возвращает все)
            
        Returns:
            DataFrame с найденными записями
        """
        self._ensure_data_loaded()
        status_lower = status.lower().strip()
        results = []
        
        categories_to_search = [category] if category and category in self.data_cache else self.categories
        
        for cat in categories_to_search:
            df = self.data_cache[cat]
            if 'Статус_тендера' in df.columns:
                mask = df['Статус_тендера'].astype(str).str.lower().str.contains(status_lower, na=False)
                found = df[mask].copy()
                found['Категория'] = cat
                results.append(found)
        
        if results:
            result_df = pd.concat(results, ignore_index=True)
            # Применяем лимит, если указан
            if limit is not None and limit > 0:
                result_df = result_df.head(limit)
            return result_df
        return pd.DataFrame()
    
    def get_top_managers_by_won_tenders(
        self, 
        category: Optional[str] = None, 
        limit: int = 10,
        by_sum: bool = False
    ) -> pd.DataFrame:
        """
        Получает топ менеджеров по выигранным тендерам.
        
        Args:
            category: Категория (если None, по всем)
            limit: Количество топ результатов
            by_sum: Если True, сортирует по сумме выигранных тендеров, иначе по количеству
            
        Returns:
            DataFrame с топ менеджерами
        """
        self._ensure_data_loaded()
        if category and category in self.data_cache:
            df = self.data_cache[category]
        elif not self.categories:
            return pd.DataFrame()
        else:
            dfs = [self.data_cache[cat] for cat in self.categories]
            df = pd.concat(dfs, ignore_index=True)
        
        if 'Менеджер' not in df.columns or 'Статус_тендера' not in df.columns:
            return pd.DataFrame()
        
        # Фильтруем только выигранные тендеры (Выиграли, Выиграли частично)
        won_statuses = ['выиграли', 'выиграли частично']
        won_mask = df['Статус_тендера'].astype(str).str.lower().str.contains('|'.join(won_statuses), na=False, regex=True)
        won_df = df[won_mask].copy()
        
        if won_df.empty:
            return pd.DataFrame()
        
        if by_sum and 'Цена подачи' in won_df.columns:
            # Считаем общую сумму по менеджерам
            # Для каждой записи: цена подачи * количество (или просто цена подачи)
            won_df['_total_price'] = won_df['Цена подачи'].fillna(0) * won_df['Количество'].fillna(1)
            manager_stats = won_df.groupby('Менеджер').agg({
                '_total_price': 'sum',
                'Чертеж': 'count'  # Количество тендеров
            }).reset_index()
            manager_stats.columns = ['Менеджер', 'Сумма_руб', 'Количество_тендеров']
            manager_stats = manager_stats.sort_values('Сумма_руб', ascending=False).head(limit)
            manager_stats['Сумма_руб'] = manager_stats['Сумма_руб'].round(2)
        else:
            # По количеству выигранных тендеров
            manager_stats = won_df['Менеджер'].value_counts().head(limit).reset_index()
            manager_stats.columns = ['Менеджер', 'Количество_тендеров']
        
        return manager_stats
    
    def get_largest_won_tender(self, category: Optional[str] = None) -> Optional[Dict[str, Any]]:
        """
        Находит самый большой выигранный тендер по сумме.
        
        Args:
            category: Категория (если None, по всем)
            
        Returns:
            Словарь с данными тендера или None
        """
        self._ensure_data_loaded()
        if category and category in self.data_cache:
            df = self.data_cache[category]
        elif not self.categories:
            return None
        else:
            dfs = [self.data_cache[cat] for cat in self.categories]
            df = pd.concat(dfs, ignore_index=True)
        
        if 'Статус_тендера' not in df.columns or 'Цена подачи' not in df.columns:
            return None
        
        # Фильтруем только выигранные тендеры
        won_statuses = ['выиграли', 'выиграли частично']
        won_mask = df['Статус_тендера'].astype(str).str.lower().str.contains('|'.join(won_statuses), na=False, regex=True)
        won_df = df[won_mask].copy()
        
        if won_df.empty:
            return None
        
        # Считаем общую сумму для каждой записи (цена * количество)
        won_df['_total_price'] = won_df['Цена подачи'].fillna(0) * won_df['Количество'].fillna(1)
        
        # Находим запись с максимальной суммой
        largest_idx = won_df['_total_price'].idxmax()
        largest = won_df.loc[largest_idx]
        
        return {
            'Номенклатура': str(largest.get('Номенклатура', '')),
            'Чертеж': str(largest.get('Чертеж', '')),
            'Заказчик': str(largest.get('Заказчик, сокр', '')),
            'Менеджер': str(largest.get('Менеджер', '')),
            'Цена_подачи': float(largest.get('Цена подачи', 0)),
            'Количество': float(largest.get('Количество', 0)),
            'Общая_сумма': float(largest['_total_price']),
            'Статус_тендера': str(largest.get('Статус_тендера', '')),
            'Дата_тендера': str(largest.get('Дата_тендера', '')),
            'Категория': str(largest.get('Категория', ''))
        }
    
    def filter_by_date_range(
        self,
        df: pd.DataFrame,
        start_date: Optional[str] = None,
        end_date: Optional[str] = None
    ) -> pd.DataFrame:
        """
        Фильтрует DataFrame по диапазону дат тендеров.
        
        Args:
            df: DataFrame для фильтрации
            start_date: Начальная дата (формат: "YYYY-MM-DD" или "DD.MM.YYYY")
            end_date: Конечная дата (формат: "YYYY-MM-DD" или "DD.MM.YYYY")
            
        Returns:
            Отфильтрованный DataFrame
        """
        self._ensure_data_loaded()
        
        if 'Дата_тендера_parsed' not in df.columns:
            logger.warning("Колонка 'Дата_тендера_parsed' отсутствует, фильтрация по датам невозможна")
            return df
        
        filtered_df = df.copy()
        
        try:
            # Преобразуем даты в datetime, если они строками
            if start_date:
                if '.' in start_date:
                    start_dt = pd.to_datetime(start_date, format='%d.%m.%Y', errors='coerce')
                else:
                    start_dt = pd.to_datetime(start_date, errors='coerce')
                if pd.notna(start_dt):
                    filtered_df = filtered_df[filtered_df['Дата_тендера_parsed'] >= start_dt]
            
            if end_date:
                if '.' in end_date:
                    end_dt = pd.to_datetime(end_date, format='%d.%m.%Y', errors='coerce')
                else:
                    end_dt = pd.to_datetime(end_date, errors='coerce')
                if pd.notna(end_dt):
                    filtered_df = filtered_df[filtered_df['Дата_тендера_parsed'] <= end_dt]
        except Exception as e:
            logger.warning(f"Ошибка при фильтрации по датам: {e}")
            return df
        
        return filtered_df
    
    def filter_results(
        self,
        df: pd.DataFrame,
        filters: Optional[Dict[str, Any]] = None
    ) -> pd.DataFrame:
        """
        Фильтрует результаты по множественным критериям.
        
        Args:
            df: DataFrame для фильтрации
            filters: Словарь с фильтрами:
                - price_min, price_max: Диапазон цен
                - weight_min, weight_max: Диапазон веса
                - customer: Фильтр по заказчику
                - material: Фильтр по материалу
                - manager: Фильтр по менеджеру
                - status: Фильтр по статусу тендера
                - date_from, date_to: Диапазон дат
                
        Returns:
            Отфильтрованный DataFrame
        """
        if filters is None or not filters:
            return df
        
        filtered_df = df.copy()
        
        try:
            # Фильтр по цене
            if 'price_min' in filters and 'Цена подачи' in filtered_df.columns:
                price_min = filters['price_min']
                if price_min is not None:
                    filtered_df = filtered_df[filtered_df['Цена подачи'] >= price_min]
            
            if 'price_max' in filters and 'Цена подачи' in filtered_df.columns:
                price_max = filters['price_max']
                if price_max is not None:
                    filtered_df = filtered_df[filtered_df['Цена подачи'] <= price_max]
            
            # Фильтр по весу
            if 'weight_min' in filters and 'ГП вес, кг' in filtered_df.columns:
                weight_min = filters['weight_min']
                if weight_min is not None:
                    filtered_df = filtered_df[filtered_df['ГП вес, кг'] >= weight_min]
            
            if 'weight_max' in filters and 'ГП вес, кг' in filtered_df.columns:
                weight_max = filters['weight_max']
                if weight_max is not None:
                    filtered_df = filtered_df[filtered_df['ГП вес, кг'] <= weight_max]
            
            # Фильтр по заказчику
            if 'customer' in filters:
                customer = filters['customer']
                if customer:
                    customer_col = 'Заказчик, сокр' if 'Заказчик, сокр' in filtered_df.columns else 'Заказчик'
                    if customer_col in filtered_df.columns:
                        filtered_df = filtered_df[
                            filtered_df[customer_col].astype(str).str.lower().str.contains(
                                customer.lower(), na=False
                            )
                        ]
            
            # Фильтр по материалу
            if 'material' in filters and 'Материал' in filtered_df.columns:
                material = filters['material']
                if material:
                    filtered_df = filtered_df[
                        filtered_df['Материал'].astype(str).str.lower().str.contains(
                            material.lower(), na=False
                        )
                    ]
            
            # Фильтр по менеджеру
            if 'manager' in filters and 'Менеджер' in filtered_df.columns:
                manager = filters['manager']
                if manager:
                    filtered_df = filtered_df[
                        filtered_df['Менеджер'].astype(str).str.lower().str.contains(
                            manager.lower(), na=False
                        )
                    ]
            
            # Фильтр по статусу тендера
            if 'status' in filters and 'Статус_тендера' in filtered_df.columns:
                status = filters['status']
                if status:
                    filtered_df = filtered_df[
                        filtered_df['Статус_тендера'].astype(str).str.lower().str.contains(
                            status.lower(), na=False
                        )
                    ]
            
            # Фильтр по датам
            if 'date_from' in filters or 'date_to' in filters:
                filtered_df = self.filter_by_date_range(
                    filtered_df,
                    start_date=filters.get('date_from'),
                    end_date=filters.get('date_to')
                )
        
        except Exception as e:
            logger.error(f"Ошибка при фильтрации результатов: {e}", exc_info=True)
            return df
        
        return filtered_df
    
    def sort_results(
        self,
        df: pd.DataFrame,
        sort_by: Optional[str] = None,
        sort_order: str = 'asc'
    ) -> pd.DataFrame:
        """
        Сортирует результаты по указанному столбцу.
        
        Args:
            df: DataFrame для сортировки
            sort_by: Название столбца для сортировки (если None, не сортирует)
            sort_order: Порядок сортировки ('asc' или 'desc')
            
        Returns:
            Отсортированный DataFrame
        """
        if df.empty or not sort_by:
            return df
        
        if sort_by not in df.columns:
            logger.warning(f"Столбец {sort_by} не найден для сортировки")
            return df
        
        try:
            ascending = sort_order.lower() in ['asc', 'ascending', 'возр', 'возрастание']
            sorted_df = df.sort_values(by=sort_by, ascending=ascending, na_position='last')
            return sorted_df
        except Exception as e:
            logger.error(f"Ошибка при сортировке результатов: {e}", exc_info=True)
            return df
    
    def compare_periods(
        self,
        category: Optional[str] = None,
        period1_start: Optional[str] = None,
        period1_end: Optional[str] = None,
        period2_start: Optional[str] = None,
        period2_end: Optional[str] = None
    ) -> Dict[str, Any]:
        """
        Сравнивает статистику двух периодов.
        
        Args:
            category: Категория (если None, по всем)
            period1_start: Начало первого периода
            period1_end: Конец первого периода
            period2_start: Начало второго периода
            period2_end: Конец второго периода
            
        Returns:
            Словарь с сравнительной статистикой
        """
        self._ensure_data_loaded()
        
        # Получаем данные
        if category and category in self.data_cache:
            df = self.data_cache[category]
        elif not self.categories:
            return {'error': 'Нет доступных данных'}
        else:
            dfs = [self.data_cache[cat] for cat in self.categories]
            df = pd.concat(dfs, ignore_index=True)
        
        # Фильтруем по периодам
        period1_df = df.copy()
        if period1_start or period1_end:
            period1_df = self.filter_by_date_range(period1_df, period1_start, period1_end)
        
        period2_df = df.copy()
        if period2_start or period2_end:
            period2_df = self.filter_by_date_range(period2_df, period2_start, period2_end)
        
        # Считаем статистику для каждого периода
        stats1 = self._calculate_period_stats(period1_df)
        stats2 = self._calculate_period_stats(period2_df)
        
        # Вычисляем изменения
        comparison = {
            'period1': stats1,
            'period2': stats2,
            'comparison': {}
        }
        
        for key in stats1:
            if isinstance(stats1[key], (int, float)) and isinstance(stats2.get(key), (int, float)):
                val1 = stats1[key]
                val2 = stats2.get(key, 0)
                if val1 != 0:
                    change_percent = ((val2 - val1) / val1) * 100
                else:
                    change_percent = 100 if val2 > 0 else 0
                comparison['comparison'][key] = {
                    'absolute_change': val2 - val1,
                    'percent_change': round(change_percent, 2)
                }
        
        return comparison
    
    def _calculate_period_stats(self, df: pd.DataFrame) -> Dict[str, Any]:
        """Вычисляет статистику для периода."""
        stats = {
            'total_records': len(df),
            'total_weight_kg': float(df['ГП вес, кг'].sum() if 'ГП вес, кг' in df.columns else 0),
            'total_price': float(df['Цена подачи'].sum() if 'Цена подачи' in df.columns else 0),
        }
        
        if 'Заказчик, сокр' in df.columns:
            stats['unique_customers'] = int(df['Заказчик, сокр'].nunique())
        elif 'Заказчик' in df.columns:
            stats['unique_customers'] = int(df['Заказчик'].nunique())
        
        if 'Материал' in df.columns:
            stats['unique_materials'] = int(df['Материал'].nunique())
        
        # Статистика по тендерам
        if 'Статус_тендера' in df.columns:
            won_statuses = ['выиграли', 'выиграли частично']
            won_mask = df['Статус_тендера'].astype(str).str.lower().str.contains('|'.join(won_statuses), na=False, regex=True)
            stats['won_tenders'] = int(won_mask.sum())
        
        return stats
    
    def get_tender_statistics(self, category: Optional[str] = None) -> Dict[str, Any]:
        """
        Получает статистику по тендерам.
        
        Args:
            category: Категория (если None, по всем)
            
        Returns:
            Словарь со статистикой по тендерам
        """
        self._ensure_data_loaded()
        if category and category in self.data_cache:
            df = self.data_cache[category]
        elif not self.categories:
            return {}
        else:
            dfs = [self.data_cache[cat] for cat in self.categories]
            df = pd.concat(dfs, ignore_index=True)
        
        if 'Статус_тендера' not in df.columns:
            return {}
        
        stats = {}
        
        # Подсчитываем по статусам
        status_counts = df['Статус_тендера'].value_counts().to_dict()
        
        # Нормализуем статусы (приводим к нижнему регистру для подсчета)
        status_normalized = df['Статус_тендера'].astype(str).str.lower().str.strip()
        
        stats['total_tenders'] = len(df[status_normalized != ''])
        stats['won'] = int((status_normalized.str.contains('выиграли', na=False)).sum())
        stats['won_partially'] = int((status_normalized.str.contains('выиграли частично', na=False)).sum())
        stats['lost'] = int((status_normalized.str.contains('проиграли', na=False)).sum())
        stats['cancelled'] = int((status_normalized.str.contains('отменен', na=False)).sum())
        stats['submitted'] = int((status_normalized.str.contains('подались', na=False)).sum())
        stats['not_participated'] = int((status_normalized.str.contains('не участвовали', na=False)).sum())
        
        # Сумма выигранных тендеров
        if 'Цена подачи' in df.columns:
            won_statuses = ['выиграли', 'выиграли частично']
            won_mask = status_normalized.str.contains('|'.join(won_statuses), na=False, regex=True)
            won_df = df[won_mask]
            if not won_df.empty:
                won_df['_total'] = won_df['Цена подачи'].fillna(0) * won_df['Количество'].fillna(1)
                stats['won_sum_rub'] = float(won_df['_total'].sum())
            else:
                stats['won_sum_rub'] = 0.0
        else:
            stats['won_sum_rub'] = 0.0
        
        stats['status_distribution'] = {k: int(v) for k, v in status_counts.items() if k and str(k).strip()}
        
        return stats
    
    def generate_chart_customers(self, df: pd.DataFrame, limit: int = 10) -> Optional[str]:
        """
        Генерирует график по заказчикам (top заказчиков по количеству).
        
        Args:
            df: DataFrame с данными
            limit: Количество топ заказчиков для отображения
            
        Returns:
            Base64 строка изображения или None при ошибке
        """
        try:
            customer_col = 'Заказчик, сокр' if 'Заказчик, сокр' in df.columns else 'Заказчик'
            if customer_col not in df.columns or df[customer_col].isna().all():
                return None
            
            top_customers = df[customer_col].value_counts().head(limit)
            
            if top_customers.empty:
                return None
            
            fig, ax = plt.subplots(figsize=(10, 6))
            top_customers.sort_values().plot(kind='barh', ax=ax, color='steelblue')
            ax.set_title(f'Топ-{limit} заказчиков по количеству заказов', fontsize=14, fontweight='bold')
            ax.set_xlabel('Количество заказов', fontsize=12)
            ax.set_ylabel('')
            plt.tight_layout()
            
            return self._fig_to_base64(fig)
        except Exception as e:
            logger.error(f"Ошибка при генерации графика по заказчикам: {e}", exc_info=True)
            return None
    
    def generate_chart_materials(self, df: pd.DataFrame, limit: int = 10) -> Optional[str]:
        """
        Генерирует график по материалам (pie chart).
        
        Args:
            df: DataFrame с данными
            limit: Количество топ материалов для отображения
            
        Returns:
            Base64 строка изображения или None при ошибке
        """
        try:
            if 'Материал' not in df.columns or df['Материал'].isna().all():
                return None
            
            material_counts = df['Материал'].value_counts().head(limit)
            
            if material_counts.empty:
                return None
            
            fig, ax = plt.subplots(figsize=(10, 6))
            ax.pie(material_counts.values, labels=material_counts.index, autopct='%1.1f%%', startangle=90)
            ax.axis('equal')
            ax.set_title(f'Популярные материалы (топ-{limit})', fontsize=14, fontweight='bold')
            plt.tight_layout()
            
            return self._fig_to_base64(fig)
        except Exception as e:
            logger.error(f"Ошибка при генерации графика по материалам: {e}", exc_info=True)
            return None
    
    def generate_chart_managers(self, df: pd.DataFrame, by_sum: bool = False, limit: int = 10) -> Optional[str]:
        """
        Генерирует график по менеджерам (топ менеджеров по количеству или сумме тендеров).
        
        Args:
            df: DataFrame с данными
            by_sum: Если True, сортирует по сумме, иначе по количеству
            limit: Количество топ менеджеров для отображения
            
        Returns:
            Base64 строка изображения или None при ошибке
        """
        try:
            if 'Менеджер' not in df.columns or df['Менеджер'].isna().all():
                return None
            
            # Фильтруем только выигранные тендеры для графика менеджеров
            if 'Статус_тендера' in df.columns:
                won_statuses = ['выиграли', 'выиграли частично']
                status_mask = df['Статус_тендера'].astype(str).str.lower().str.contains('|'.join(won_statuses), na=False, regex=True)
                manager_df = df[status_mask].copy()
            else:
                manager_df = df.copy()
            
            if manager_df.empty:
                return None
            
            if by_sum and 'Цена подачи' in manager_df.columns:
                # По сумме выигранных тендеров
                manager_df['_total'] = manager_df['Цена подачи'].fillna(0) * manager_df['Количество'].fillna(1)
                manager_stats = manager_df.groupby('Менеджер')['_total'].sum().sort_values(ascending=False).head(limit)
                
                fig, ax = plt.subplots(figsize=(10, 6))
                manager_stats.plot(kind='barh', ax=ax, color='seagreen')
                ax.set_title(f'Топ-{limit} менеджеров по сумме выигранных тендеров', fontsize=14, fontweight='bold')
                ax.set_xlabel('Сумма (руб)', fontsize=12)
                ax.set_ylabel('')
                
                # Форматируем значения на оси X
                ax.xaxis.set_major_formatter(plt.FuncFormatter(lambda x, p: f'{x:,.0f}'.replace(',', ' ')))
            else:
                # По количеству тендеров
                manager_stats = manager_df['Менеджер'].value_counts().head(limit)
                
                fig, ax = plt.subplots(figsize=(10, 6))
                manager_stats.sort_values().plot(kind='barh', ax=ax, color='orange')
                ax.set_title(f'Топ-{limit} менеджеров по количеству выигранных тендеров', fontsize=14, fontweight='bold')
                ax.set_xlabel('Количество тендеров', fontsize=12)
                ax.set_ylabel('')
            
            plt.tight_layout()
            
            return self._fig_to_base64(fig)
        except Exception as e:
            logger.error(f"Ошибка при генерации графика по менеджерам: {e}", exc_info=True)
            return None
    
    def generate_chart_prices(self, df: pd.DataFrame) -> Optional[str]:
        """
        Генерирует график по ценам (средняя цена популярных материалов).
        
        Args:
            df: DataFrame с данными
            
        Returns:
            Base64 строка изображения или None при ошибке
        """
        try:
            if 'Материал' not in df.columns or 'Цена подачи за кг' not in df.columns:
                return None
            
            if df['Материал'].isna().all() or df['Цена подачи за кг'].isna().all():
                return None
            
            # Топ материалов по количеству
            top_materials = df['Материал'].value_counts().head(10).index.tolist()
            
            # Средняя цена для топ материалов
            avg_price = (
                df[df['Материал'].isin(top_materials)]
                .groupby('Материал')['Цена подачи за кг']
                .mean()
                .reindex(top_materials)
            )
            
            if avg_price.empty:
                return None
            
            fig, ax = plt.subplots(figsize=(10, 6))
            avg_price.plot(kind='bar', ax=ax, color='slateblue')
            ax.set_title('Средняя цена за кг по популярным материалам', fontsize=14, fontweight='bold')
            ax.set_ylabel('Цена за кг (руб)', fontsize=12)
            ax.tick_params(axis='x', rotation=45)
            if ax.get_legend():
                ax.get_legend().remove()
            
            # Добавляем значения на столбцы
            for idx, value in enumerate(avg_price.values):
                if pd.notna(value):
                    ax.text(idx, value, f"{value:,.0f}".replace(',', ' '), 
                           ha='center', va='bottom', fontsize=10, fontweight='600')
            
            plt.tight_layout()
            
            return self._fig_to_base64(fig)
        except Exception as e:
            logger.error(f"Ошибка при генерации графика по ценам: {e}", exc_info=True)
            return None
    
    def generate_chart_timeline(
        self,
        df: pd.DataFrame,
        date_column: str = 'Дата_тендера_parsed',
        value_column: Optional[str] = None,
        group_by: str = 'month',
        limit: Optional[int] = None
    ) -> Optional[str]:
        """
        Генерирует график временных рядов (динамика по месяцам/кварталам).
        
        Args:
            df: DataFrame с данными
            date_column: Колонка с датами
            value_column: Колонка со значениями для агрегации (если None, считает количество)
            group_by: Группировка ('month', 'quarter', 'year')
            limit: Ограничение количества точек (если None, все)
            
        Returns:
            Base64 строка изображения или None при ошибке
        """
        try:
            if date_column not in df.columns:
                logger.warning(f"Колонка {date_column} отсутствует для временного ряда")
                return None
            
            # Фильтруем только записи с валидными датами
            df_with_dates = df[df[date_column].notna()].copy()
            if df_with_dates.empty:
                return None
            
            # Группируем по периодам
            df_with_dates['_period'] = pd.to_datetime(df_with_dates[date_column])
            
            if group_by == 'month':
                df_with_dates['_period_key'] = df_with_dates['_period'].dt.to_period('M')
                period_format = '%Y-%m'
            elif group_by == 'quarter':
                df_with_dates['_period_key'] = df_with_dates['_period'].dt.to_period('Q')
                period_format = '%Y-Q%q'
            elif group_by == 'year':
                df_with_dates['_period_key'] = df_with_dates['_period'].dt.to_period('Y')
                period_format = '%Y'
            else:
                df_with_dates['_period_key'] = df_with_dates['_period'].dt.to_period('M')
                period_format = '%Y-%m'
            
            # Агрегируем данные
            if value_column and value_column in df_with_dates.columns:
                # Суммируем значения
                timeline_data = df_with_dates.groupby('_period_key')[value_column].sum().sort_index()
            else:
                # Считаем количество
                timeline_data = df_with_dates.groupby('_period_key').size().sort_index()
            
            if timeline_data.empty:
                return None
            
            # Применяем лимит если указан
            if limit and len(timeline_data) > limit:
                timeline_data = timeline_data.tail(limit)
            
            # Конвертируем периоды в строки для отображения
            timeline_data.index = timeline_data.index.astype(str)
            
            fig, ax = plt.subplots(figsize=(12, 6))
            timeline_data.plot(kind='line', ax=ax, marker='o', linewidth=2, markersize=6, color='steelblue')
            ax.set_title(f'Динамика тендеров по {group_by}', fontsize=14, fontweight='bold')
            ax.set_xlabel('Период', fontsize=12)
            ax.set_ylabel('Количество' if not value_column else f'Сумма {value_column}', fontsize=12)
            ax.grid(True, alpha=0.3)
            plt.xticks(rotation=45, ha='right')
            plt.tight_layout()
            
            return self._fig_to_base64(fig)
        except Exception as e:
            logger.error(f"Ошибка при генерации временного ряда: {e}", exc_info=True)
            return None
    
    def generate_chart_scatter(
        self,
        df: pd.DataFrame,
        x_column: str = 'ГП вес, кг',
        y_column: str = 'Цена подачи',
        color_by: Optional[str] = None,
        limit: int = 500
    ) -> Optional[str]:
        """
        Генерирует scatter plot для корреляций (например, вес vs цена).
        
        Args:
            df: DataFrame с данными
            x_column: Колонка для оси X
            y_column: Колонка для оси Y
            color_by: Колонка для цветовой кодировки (опционально)
            limit: Максимальное количество точек (для производительности)
            
        Returns:
            Base64 строка изображения или None при ошибке
        """
        try:
            if x_column not in df.columns or y_column not in df.columns:
                return None
            
            # Фильтруем только валидные данные
            scatter_df = df[[x_column, y_column]].dropna()
            if color_by and color_by in df.columns:
                scatter_df[color_by] = df[color_by]
            
            if scatter_df.empty:
                return None
            
            # Ограничиваем количество точек для производительности
            if len(scatter_df) > limit:
                scatter_df = scatter_df.sample(n=limit, random_state=42)
            
            fig, ax = plt.subplots(figsize=(10, 6))
            
            if color_by and color_by in scatter_df.columns:
                # Scatter plot с цветовой кодировкой
                unique_values = scatter_df[color_by].nunique()
                if unique_values <= 10:
                    # Используем дискретные цвета для небольшого количества категорий
                    scatter_df.plot.scatter(
                        x=x_column,
                        y=y_column,
                        c=scatter_df[color_by].astype('category').cat.codes,
                        colormap='viridis',
                        ax=ax,
                        alpha=0.6,
                        s=50
                    )
                else:
                    # Для большого количества категорий используем один цвет
                    ax.scatter(
                        scatter_df[x_column],
                        scatter_df[y_column],
                        alpha=0.6,
                        s=50,
                        c='steelblue'
                    )
            else:
                # Обычный scatter plot
                ax.scatter(
                    scatter_df[x_column],
                    scatter_df[y_column],
                    alpha=0.6,
                    s=50,
                    c='steelblue'
                )
            
            ax.set_title(f'Корреляция: {x_column} vs {y_column}', fontsize=14, fontweight='bold')
            ax.set_xlabel(x_column, fontsize=12)
            ax.set_ylabel(y_column, fontsize=12)
            ax.grid(True, alpha=0.3)
            plt.tight_layout()
            
            return self._fig_to_base64(fig)
        except Exception as e:
            logger.error(f"Ошибка при генерации scatter plot: {e}", exc_info=True)
            return None
    
    def generate_chart_heatmap(
        self,
        df: pd.DataFrame,
        x_column: str = 'Заказчик, сокр',
        y_column: str = 'Менеджер',
        value_column: str = 'Цена подачи',
        aggregation: str = 'sum',
        limit_x: int = 15,
        limit_y: int = 15
    ) -> Optional[str]:
        """
        Генерирует heatmap для многомерных данных (корреляции заказчик-менеджер).
        
        Args:
            df: DataFrame с данными
            x_column: Колонка для оси X
            y_column: Колонка для оси Y
            value_column: Колонка со значениями для агрегации
            aggregation: Тип агрегации ('sum', 'mean', 'count')
            limit_x: Максимальное количество значений по оси X
            limit_y: Максимальное количество значений по оси Y
            
        Returns:
            Base64 строка изображения или None при ошибке
        """
        try:
            required_cols = [x_column, y_column]
            if value_column:
                required_cols.append(value_column)
            
            missing_cols = [col for col in required_cols if col not in df.columns]
            if missing_cols:
                logger.warning(f"Отсутствуют колонки для heatmap: {missing_cols}")
                return None
            
            # Фильтруем валидные данные
            heatmap_df = df[required_cols].dropna()
            if heatmap_df.empty:
                return None
            
            # Берем топ значения для каждой оси
            top_x = heatmap_df[x_column].value_counts().head(limit_x).index.tolist()
            top_y = heatmap_df[y_column].value_counts().head(limit_y).index.tolist()
            
            # Фильтруем данные
            heatmap_df = heatmap_df[
                heatmap_df[x_column].isin(top_x) & 
                heatmap_df[y_column].isin(top_y)
            ]
            
            if heatmap_df.empty:
                return None
            
            # Создаем сводную таблицу
            if value_column:
                if aggregation == 'sum':
                    pivot = heatmap_df.pivot_table(
                        values=value_column,
                        index=y_column,
                        columns=x_column,
                        aggfunc='sum',
                        fill_value=0
                    )
                elif aggregation == 'mean':
                    pivot = heatmap_df.pivot_table(
                        values=value_column,
                        index=y_column,
                        columns=x_column,
                        aggfunc='mean',
                        fill_value=0
                    )
                else:
                    pivot = heatmap_df.pivot_table(
                        values=value_column,
                        index=y_column,
                        columns=x_column,
                        aggfunc='count',
                        fill_value=0
                    )
            else:
                pivot = heatmap_df.pivot_table(
                    index=y_column,
                    columns=x_column,
                    aggfunc='size',
                    fill_value=0
                )
            
            if pivot.empty:
                return None
            
            # Генерируем heatmap
            fig, ax = plt.subplots(figsize=(max(12, len(pivot.columns) * 0.8), max(8, len(pivot.index) * 0.5)))
            
            # Используем seaborn для более красивого heatmap
            sns.heatmap(
                pivot,
                annot=True,
                fmt='.0f',
                cmap='YlOrRd',
                ax=ax,
                cbar_kws={'label': value_column if value_column else 'Количество'},
                linewidths=0.5,
                linecolor='gray'
            )
            
            ax.set_title(
                f'Heatmap: {y_column} vs {x_column} ({aggregation})',
                fontsize=14,
                fontweight='bold'
            )
            ax.set_xlabel(x_column, fontsize=11)
            ax.set_ylabel(y_column, fontsize=11)
            plt.xticks(rotation=45, ha='right')
            plt.yticks(rotation=0)
            plt.tight_layout()
            
            return self._fig_to_base64(fig)
        except Exception as e:
            logger.error(f"Ошибка при генерации heatmap: {e}", exc_info=True)
            return None
    
    def compare_periods(
        self,
        period1_start: str,
        period1_end: str,
        period2_start: str,
        period2_end: str,
        category: Optional[str] = None,
        metric: str = 'count'  # 'count', 'sum_price', 'sum_weight'
    ) -> Dict[str, Any]:
        """
        Сравнивает два периода по различным метрикам.
        
        Args:
            period1_start: Начало первого периода
            period1_end: Конец первого периода
            period2_start: Начало второго периода
            period2_end: Конец второго периода
            category: Категория для сравнения (если None, все категории)
            metric: Метрика для сравнения ('count', 'sum_price', 'sum_weight')
            
        Returns:
            Словарь со сравнением периодов
        """
        self._ensure_data_loaded()
        
        try:
            # Получаем данные
            if category and category in self.data_cache:
                df = self.data_cache[category].copy()
            else:
                dfs = [self.data_cache[cat] for cat in self.categories if cat in self.data_cache]
                df = pd.concat(dfs, ignore_index=True) if dfs else pd.DataFrame()
            
            if df.empty:
                return {
                    'error': 'Нет данных для сравнения',
                    'period1': {'value': 0},
                    'period2': {'value': 0},
                    'change': {'absolute': 0, 'percentage': 0}
                }
            
            # Фильтруем данные по периодам
            period1_df = self.filter_by_date_range(df, period1_start, period1_end)
            period2_df = self.filter_by_date_range(df, period2_start, period2_end)
            
            # Вычисляем метрики
            def calculate_metric(df_subset: pd.DataFrame, metric_name: str) -> float:
                if metric_name == 'count':
                    return float(len(df_subset))
                elif metric_name == 'sum_price' and 'Цена подачи' in df_subset.columns:
                    return float(df_subset['Цена подачи'].fillna(0).sum())
                elif metric_name == 'sum_weight' and 'ГП вес, кг' in df_subset.columns:
                    return float(df_subset['ГП вес, кг'].fillna(0).sum())
                return 0.0
            
            period1_value = calculate_metric(period1_df, metric)
            period2_value = calculate_metric(period2_df, metric)
            
            # Вычисляем изменение
            change_absolute = period2_value - period1_value
            change_percentage = ((period2_value - period1_value) / period1_value * 100) if period1_value > 0 else 0
            
            return {
                'period1': {
                    'start': period1_start,
                    'end': period1_end,
                    'value': period1_value,
                    'records_count': len(period1_df)
                },
                'period2': {
                    'start': period2_start,
                    'end': period2_end,
                    'value': period2_value,
                    'records_count': len(period2_df)
                },
                'change': {
                    'absolute': change_absolute,
                    'percentage': change_percentage,
                    'is_increase': change_absolute > 0
                },
                'metric': metric,
                'category': category or 'all'
            }
        
        except Exception as e:
            logger.error(f"Ошибка при сравнении периодов: {e}", exc_info=True)
            return {
                'error': str(e),
                'period1': {'value': 0},
                'period2': {'value': 0},
                'change': {'absolute': 0, 'percentage': 0}
            }
    
    def compare_categories(
        self,
        category1: str,
        category2: str,
        metric: str = 'count'  # 'count', 'sum_price', 'sum_weight'
    ) -> Dict[str, Any]:
        """
        Сравнивает две категории по различным метрикам.
        
        Args:
            category1: Первая категория
            category2: Вторая категория
            metric: Метрика для сравнения ('count', 'sum_price', 'sum_weight')
            
        Returns:
            Словарь со сравнением категорий
        """
        self._ensure_data_loaded()
        
        try:
            # Проверяем наличие категорий
            if category1 not in self.data_cache or category2 not in self.data_cache:
                missing = [c for c in [category1, category2] if c not in self.data_cache]
                return {
                    'error': f'Категории не найдены: {", ".join(missing)}',
                    'category1': {'value': 0},
                    'category2': {'value': 0},
                    'change': {'absolute': 0, 'percentage': 0}
                }
            
            df1 = self.data_cache[category1].copy()
            df2 = self.data_cache[category2].copy()
            
            # Вычисляем метрики
            def calculate_metric(df_subset: pd.DataFrame, metric_name: str) -> float:
                if metric_name == 'count':
                    return float(len(df_subset))
                elif metric_name == 'sum_price' and 'Цена подачи' in df_subset.columns:
                    return float(df_subset['Цена подачи'].fillna(0).sum())
                elif metric_name == 'sum_weight' and 'ГП вес, кг' in df_subset.columns:
                    return float(df_subset['ГП вес, кг'].fillna(0).sum())
                return 0.0
            
            category1_value = calculate_metric(df1, metric)
            category2_value = calculate_metric(df2, metric)
            
            # Вычисляем изменение
            change_absolute = category2_value - category1_value
            change_percentage = ((category2_value - category1_value) / category1_value * 100) if category1_value > 0 else 0
            
            return {
                'category1': {
                    'name': category1,
                    'value': category1_value,
                    'records_count': len(df1)
                },
                'category2': {
                    'name': category2,
                    'value': category2_value,
                    'records_count': len(df2)
                },
                'change': {
                    'absolute': change_absolute,
                    'percentage': change_percentage,
                    'is_increase': change_absolute > 0
                },
                'metric': metric
            }
        
        except Exception as e:
            logger.error(f"Ошибка при сравнении категорий: {e}", exc_info=True)
            return {
                'error': str(e),
                'category1': {'value': 0},
                'category2': {'value': 0},
                'change': {'absolute': 0, 'percentage': 0}
            }
    
    def generate_comparison_chart(
        self,
        comparison_data: Dict[str, Any],
        chart_type: str = 'bar'  # 'bar', 'line'
    ) -> Optional[str]:
        """
        Генерирует график для сравнения двух периодов или категорий.
        
        Args:
            comparison_data: Данные сравнения из compare_periods или compare_categories
            chart_type: Тип графика ('bar' или 'line')
            
        Returns:
            Base64 строка изображения или None при ошибке
        """
        try:
            if 'error' in comparison_data:
                return None
            
            fig, ax = plt.subplots(figsize=(10, 6))
            
            # Определяем названия для осей
            if 'period1' in comparison_data:
                # Сравнение периодов
                labels = [
                    f"{comparison_data['period1']['start']} - {comparison_data['period1']['end']}",
                    f"{comparison_data['period2']['start']} - {comparison_data['period2']['end']}"
                ]
                values = [
                    comparison_data['period1']['value'],
                    comparison_data['period2']['value']
                ]
                title = f"Сравнение периодов: {comparison_data['metric']}"
            else:
                # Сравнение категорий
                labels = [
                    comparison_data['category1']['name'],
                    comparison_data['category2']['name']
                ]
                values = [
                    comparison_data['category1']['value'],
                    comparison_data['category2']['value']
                ]
                title = f"Сравнение категорий: {comparison_data['metric']}"
            
            if chart_type == 'bar':
                bars = ax.bar(labels, values, color=['steelblue', 'seagreen'], alpha=0.7)
                ax.set_ylabel('Значение', fontsize=12)
            else:
                ax.plot(labels, values, marker='o', linewidth=2, markersize=8, color='steelblue')
                ax.set_ylabel('Значение', fontsize=12)
                ax.grid(True, alpha=0.3)
            
            ax.set_title(title, fontsize=14, fontweight='bold')
            ax.set_xlabel('', fontsize=12)
            
            # Добавляем значения на график
            for i, (label, value) in enumerate(zip(labels, values)):
                ax.text(i, value, f'{value:,.0f}'.replace(',', ' '), 
                       ha='center', va='bottom', fontsize=10)
            
            plt.xticks(rotation=45, ha='right')
            plt.tight_layout()
            
            return self._fig_to_base64(fig)
        
        except Exception as e:
            logger.error(f"Ошибка при генерации графика сравнения: {e}", exc_info=True)
            return None
    
    def _fig_to_base64(self, fig: plt.Figure) -> str:
        """
        Конвертирует matplotlib figure в base64 строку.
        
        Args:
            fig: matplotlib Figure
            
        Returns:
            Base64 строка изображения
        """
        buffer = BytesIO()
        fig.savefig(buffer, format='png', bbox_inches='tight', dpi=100)
        plt.close(fig)
        buffer.seek(0)
        return base64.b64encode(buffer.read()).decode('utf-8')
    
    def format_search_result(
        self,
        results: pd.DataFrame,
        include_charts: bool = False,
        max_records: int = 20,
        offset: int = 0
    ) -> Dict[str, Any]:
        """
        Форматирует результат поиска для вывода с поддержкой пагинации.
        
        Args:
            results: DataFrame с результатами поиска
            include_charts: Включать ли графики
            max_records: Максимальное количество записей для отображения
            offset: Смещение для пагинации (сколько записей пропустить)
            
        Returns:
            Словарь с отформатированным результатом
        """
        if results.empty:
            # Генерируем дружелюбное сообщение с предложениями
            empty_message = SuggestionsGenerator.generate_empty_result_message(
                'search',
                query=None,  # Можно передать query если будет доступен
                category=None  # Можно передать category если будет доступен
            )
            return {
                'text': empty_message,
                'records': [],
                'total_count': 0,
                'shown_count': 0,
                'has_more': False,
                'charts': []
            }
        
        total_count = len(results)
        
        # Применяем пагинацию
        start_idx = offset
        end_idx = offset + max_records
        display_df = results.iloc[start_idx:end_idx]
        shown_count = len(display_df)
        has_more = end_idx < total_count
        
        # Формируем текстовый ответ с использованием форматтера
        summary = ResponseFormatter.format_summary(total_count, shown_count, "записей")
        text_parts = [summary]
        
        if total_count > max_records:
            if offset > 0:
                text_parts.append(f"Показаны записи {offset + 1}-{offset + shown_count} из {total_count}:")
            else:
                text_parts.append(f"Показаны первые {shown_count} записей:")
        else:
            text_parts.append(f"Показаны все {shown_count} записей:")
        
        text_parts.append("")
        
        # Форматируем записи с использованием форматтера
        records = []
        for idx, row in display_df.iterrows():
            record = {}
            for col in display_df.columns:
                value = row[col]
                if pd.notna(value):
                    # Форматируем числовые значения с использованием форматтера
                    if isinstance(value, (int, float)):
                        if 'цена' in col.lower() or 'Цена' in col or 'price' in col.lower():
                            record[col] = ResponseFormatter.format_price(value)
                        elif 'вес' in col.lower() or 'Вес' in col or 'weight' in col.lower():
                            record[col] = ResponseFormatter.format_weight(value)
                        elif 'Количество' in col or 'количество' in col.lower() or 'count' in col.lower():
                            record[col] = ResponseFormatter.format_count(value)
                        else:
                            record[col] = ResponseFormatter.format_number(value)
                    else:
                        record[col] = str(value)
                else:
                    record[col] = ''
            records.append(record)
        
        # Генерируем графики если нужно
        charts = []
        if include_charts and not results.empty:
            # График по заказчикам
            customer_chart = self.generate_chart_customers(results)
            if customer_chart:
                charts.append({
                    'title': 'Топ заказчиков',
                    'image': f"data:image/png;base64,{customer_chart}"
                })
            
            # График по материалам
            material_chart = self.generate_chart_materials(results)
            if material_chart:
                charts.append({
                    'title': 'Популярные материалы',
                    'image': f"data:image/png;base64,{material_chart}"
                })
            
            # График по ценам (только если есть данные)
            if 'Цена подачи за кг' in results.columns and results['Цена подачи за кг'].notna().any():
                price_chart = self.generate_chart_prices(results)
                if price_chart:
                    charts.append({
                        'title': 'Средняя цена по материалам',
                        'image': f"data:image/png;base64,{price_chart}"
                    })
        
        return {
            'text': '\n'.join(text_parts),
            'records': records,
            'total_count': total_count,
            'shown_count': shown_count,
            'offset': offset,
            'has_more': has_more,
            'next_offset': offset + shown_count if has_more else None,
            'charts': charts
        }
