"""Тесты для модуля кеширования app/core/cache.py."""

import pytest

from app.core.cache import (
    CacheManager,
    cached_dataset,
    get_cache_manager,
    invalidate_dataset_cache,
)


class TestCacheManager:
    """Тесты для класса CacheManager."""

    def test_cache_manager_init(self):
        """Проверяет инициализацию CacheManager."""
        manager = CacheManager()
        assert manager._cache == {}

    def test_cache_manager_set_get(self):
        """Проверяет установку и получение значений из кеша."""
        manager = CacheManager()
        manager.set('key1', 'value1')
        manager.set('key2', {'nested': 'value'})
        
        assert manager.get('key1') == 'value1'
        assert manager.get('key2') == {'nested': 'value'}
        assert manager.get('nonexistent') is None

    def test_cache_manager_invalidate_key(self):
        """Проверяет инвалидацию конкретного ключа."""
        manager = CacheManager()
        manager.set('key1', 'value1')
        manager.set('key2', 'value2')
        
        manager.invalidate('key1')
        assert manager.get('key1') is None
        assert manager.get('key2') == 'value2'

    def test_cache_manager_invalidate_all(self):
        """Проверяет очистку всего кеша."""
        manager = CacheManager()
        manager.set('key1', 'value1')
        manager.set('key2', 'value2')
        
        manager.invalidate()
        assert manager.get('key1') is None
        assert manager.get('key2') is None

    def test_cache_manager_clear_all(self):
        """Проверяет метод clear_all."""
        manager = CacheManager()
        manager.set('key1', 'value1')
        manager.set('key2', 'value2')
        
        manager.clear_all()
        assert manager.get('key1') is None
        assert manager.get('key2') is None


class TestGetCacheManager:
    """Тесты для функции get_cache_manager."""

    def test_get_cache_manager_returns_instance(self):
        """Проверяет, что get_cache_manager возвращает экземпляр CacheManager."""
        manager = get_cache_manager()
        assert isinstance(manager, CacheManager)


class TestCachedDataset:
    """Тесты для декоратора cached_dataset."""

    def test_cached_dataset_caches_result(self):
        """Проверяет, что декоратор кеширует результат функции."""
        call_count = 0

        @cached_dataset(maxsize=1)
        def test_function(x: int) -> int:
            nonlocal call_count
            call_count += 1
            return x * 2

        # Первый вызов
        result1 = test_function(5)
        assert result1 == 10
        assert call_count == 1

        # Второй вызов с теми же параметрами - должен использовать кеш
        result2 = test_function(5)
        assert result2 == 10
        assert call_count == 1  # Функция не вызывалась снова

        # Вызов с другими параметрами
        result3 = test_function(10)
        assert result3 == 20
        assert call_count == 2

    def test_cached_dataset_cache_clear(self):
        """Проверяет, что можно очистить кеш функции."""
        call_count = 0

        @cached_dataset(maxsize=1)
        def test_function(x: int) -> int:
            nonlocal call_count
            call_count += 1
            return x * 2

        test_function(5)
        assert call_count == 1

        test_function(5)
        assert call_count == 1  # Из кеша

        test_function.cache_clear()  # type: ignore

        test_function(5)
        assert call_count == 2  # Функция вызвалась снова


class TestInvalidateDatasetCache:
    """Тесты для функции invalidate_dataset_cache."""

    def test_invalidate_dataset_cache_all(self):
        """Проверяет инвалидацию всего кеша."""
        @cached_dataset(maxsize=1)
        def test_function(x: int) -> int:
            return x * 2

        test_function(5)
        invalidate_dataset_cache()
        # После инвалидации функция должна вызваться снова
        test_function(5)

