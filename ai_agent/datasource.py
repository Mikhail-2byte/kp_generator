"""
Абстракция для источников данных (CSV, БД).
Позволяет легко переключаться между CSV и базой данных.
"""

import logging
from abc import ABC, abstractmethod
from typing import Any, Dict, List, Optional

import pandas as pd

logger = logging.getLogger(__name__)


class DataSource(ABC):
    """Абстрактный базовый класс для источников данных."""
    
    @abstractmethod
    def load_data(self, category: Optional[str] = None) -> pd.DataFrame:
        """
        Загружает данные для категории.
        
        Args:
            category: Категория данных (если None, загружает все)
            
        Returns:
            DataFrame с данными
        """
        pass
    
    @abstractmethod
    def get_categories(self) -> List[str]:
        """
        Возвращает список доступных категорий.
        
        Returns:
            Список категорий
        """
        pass
    
    @abstractmethod
    def search(
        self,
        query: str,
        category: Optional[str] = None,
        filters: Optional[Dict[str, Any]] = None,
        limit: Optional[int] = None
    ) -> pd.DataFrame:
        """
        Выполняет поиск по данным.
        
        Args:
            query: Поисковый запрос
            category: Категория для поиска
            filters: Дополнительные фильтры
            limit: Максимальное количество результатов
            
        Returns:
            DataFrame с результатами поиска
        """
        pass


class CSVDataSource(DataSource):
    """Источник данных из CSV файлов."""
    
    def __init__(self, data_cache: Dict[str, pd.DataFrame], categories: List[str]):
        """
        Инициализирует CSV источник данных.
        
        Args:
            data_cache: Кеш данных (словарь категория -> DataFrame)
            categories: Список категорий
        """
        self.data_cache = data_cache
        self.categories = categories
    
    def load_data(self, category: Optional[str] = None) -> pd.DataFrame:
        """Загружает данные из CSV."""
        if category:
            if category in self.data_cache:
                return self.data_cache[category].copy()
            return pd.DataFrame()
        else:
            # Объединяем все категории
            dfs = [self.data_cache[cat] for cat in self.categories if cat in self.data_cache]
            return pd.concat(dfs, ignore_index=True) if dfs else pd.DataFrame()
    
    def get_categories(self) -> List[str]:
        """Возвращает список категорий."""
        return self.categories.copy()
    
    def search(
        self,
        query: str,
        category: Optional[str] = None,
        filters: Optional[Dict[str, Any]] = None,
        limit: Optional[int] = None
    ) -> pd.DataFrame:
        """Выполняет поиск в CSV данных."""
        df = self.load_data(category)
        
        if df.empty:
            return df
        
        # Простой поиск по всем текстовым колонкам
        query_lower = query.lower()
        mask = pd.Series([False] * len(df))
        
        for col in df.columns:
            if df[col].dtype == 'object':  # Текстовые колонки
                mask |= df[col].astype(str).str.lower().str.contains(query_lower, na=False)
        
        results = df[mask].copy()
        
        # Применяем фильтры если есть
        if filters:
            # Простая фильтрация по числовым полям
            if 'price_min' in filters and 'Цена подачи' in results.columns:
                results = results[results['Цена подачи'] >= filters['price_min']]
            if 'price_max' in filters and 'Цена подачи' in results.columns:
                results = results[results['Цена подачи'] <= filters['price_max']]
        
        # Применяем лимит
        if limit:
            results = results.head(limit)
        
        return results


class DatabaseDataSource(DataSource):
    """Источник данных из базы данных (для будущей реализации)."""
    
    def __init__(self, connection_string: str):
        """
        Инициализирует источник данных из БД.
        
        Args:
            connection_string: Строка подключения к БД
        """
        self.connection_string = connection_string
        self._engine = None
    
    def _get_engine(self):
        """Получает SQLAlchemy engine (ленивая инициализация)."""
        if self._engine is None:
            try:
                from sqlalchemy import create_engine
                self._engine = create_engine(self.connection_string)
            except ImportError:
                logger.error("SQLAlchemy не установлен. Установите: pip install sqlalchemy")
                raise
        return self._engine
    
    def load_data(self, category: Optional[str] = None) -> pd.DataFrame:
        """Загружает данные из БД."""
        try:
            engine = self._get_engine()
            query = "SELECT * FROM tenders"
            if category:
                query += f" WHERE category = '{category}'"
            
            return pd.read_sql(query, engine)
        except Exception as e:
            logger.error(f"Ошибка при загрузке данных из БД: {e}")
            return pd.DataFrame()
    
    def get_categories(self) -> List[str]:
        """Возвращает список категорий из БД."""
        try:
            engine = self._get_engine()
            query = "SELECT DISTINCT category FROM tenders"
            df = pd.read_sql(query, engine)
            return df['category'].tolist() if not df.empty else []
        except Exception as e:
            logger.error(f"Ошибка при получении категорий из БД: {e}")
            return []
    
    def search(
        self,
        query: str,
        category: Optional[str] = None,
        filters: Optional[Dict[str, Any]] = None,
        limit: Optional[int] = None
    ) -> pd.DataFrame:
        """Выполняет поиск в БД."""
        try:
            engine = self._get_engine()
            
            # Строим SQL запрос
            where_clauses = []
            params = {}
            
            # Поиск по текстовым полям
            search_fields = ['nomenclature', 'drawing', 'customer', 'material']
            search_conditions = []
            for field in search_fields:
                search_conditions.append(f"{field} LIKE :query")
            params['query'] = f"%{query}%"
            
            if search_conditions:
                where_clauses.append(f"({' OR '.join(search_conditions)})")
            
            # Фильтр по категории
            if category:
                where_clauses.append("category = :category")
                params['category'] = category
            
            # Дополнительные фильтры
            if filters:
                if 'price_min' in filters:
                    where_clauses.append("price >= :price_min")
                    params['price_min'] = filters['price_min']
                if 'price_max' in filters:
                    where_clauses.append("price <= :price_max")
                    params['price_max'] = filters['price_max']
            
            # Строим финальный запрос
            sql = "SELECT * FROM tenders"
            if where_clauses:
                sql += " WHERE " + " AND ".join(where_clauses)
            if limit:
                sql += f" LIMIT {limit}"
            
            return pd.read_sql(sql, engine, params=params)
        except Exception as e:
            logger.error(f"Ошибка при поиске в БД: {e}")
            return pd.DataFrame()


class DataSourceFactory:
    """Фабрика для создания источников данных."""
    
    @staticmethod
    def create_csv_source(
        data_cache: Dict[str, pd.DataFrame],
        categories: List[str]
    ) -> CSVDataSource:
        """
        Создает CSV источник данных.
        
        Args:
            data_cache: Кеш данных
            categories: Список категорий
            
        Returns:
            CSVDataSource
        """
        return CSVDataSource(data_cache, categories)
    
    @staticmethod
    def create_database_source(connection_string: str) -> DatabaseDataSource:
        """
        Создает источник данных из БД.
        
        Args:
            connection_string: Строка подключения к БД
            
        Returns:
            DatabaseDataSource
        """
        return DatabaseDataSource(connection_string)
    
    @staticmethod
    def create_source(
        source_type: str = 'csv',
        data_cache: Optional[Dict[str, pd.DataFrame]] = None,
        categories: Optional[List[str]] = None,
        connection_string: Optional[str] = None
    ) -> DataSource:
        """
        Создает источник данных по типу.
        
        Args:
            source_type: Тип источника ('csv' или 'database')
            data_cache: Кеш данных (для CSV)
            categories: Список категорий (для CSV)
            connection_string: Строка подключения (для БД)
            
        Returns:
            DataSource
        """
        if source_type == 'csv':
            if data_cache is None or categories is None:
                raise ValueError("Для CSV источника требуются data_cache и categories")
            return DataSourceFactory.create_csv_source(data_cache, categories)
        elif source_type == 'database':
            if connection_string is None:
                raise ValueError("Для БД источника требуется connection_string")
            return DataSourceFactory.create_database_source(connection_string)
        else:
            raise ValueError(f"Неизвестный тип источника: {source_type}")

