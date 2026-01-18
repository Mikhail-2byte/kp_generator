"""
Модуль для сбора метрик производительности агента.
Отслеживает время выполнения, использование памяти, счетчики запросов.
"""

import logging
import time
import tracemalloc
from collections import defaultdict
from dataclasses import dataclass, field
from datetime import datetime
from typing import Any, Dict, List, Optional

logger = logging.getLogger(__name__)


@dataclass
class RequestMetrics:
    """Метрики одного запроса."""
    request_type: str  # 'logistics', 'materials', 'duty', 'analytics', 'question'
    start_time: float
    end_time: Optional[float] = None
    duration_ms: Optional[float] = None
    memory_peak_mb: Optional[float] = None
    success: bool = True
    error_type: Optional[str] = None
    cache_hit: bool = False
    records_count: int = 0


class MetricsCollector:
    """Класс для сбора метрик производительности."""
    
    def __init__(self):
        """Инициализирует сборщик метрик."""
        self.metrics_history: List[RequestMetrics] = []
        self.request_counters: Dict[str, int] = defaultdict(int)
        self.error_counters: Dict[str, int] = defaultdict(int)
        self.cache_hit_counters: Dict[str, int] = defaultdict(int)
        self.total_duration_by_type: Dict[str, float] = defaultdict(float)
        self.memory_tracking_enabled = False
        
        # Включаем отслеживание памяти при первом использовании
        try:
            tracemalloc.start()
            self.memory_tracking_enabled = True
        except Exception as e:
            logger.warning(f"Не удалось включить отслеживание памяти: {e}")
    
    def start_request(self, request_type: str) -> RequestMetrics:
        """
        Начинает отслеживание метрик для запроса.
        
        Args:
            request_type: Тип запроса
            
        Returns:
            Объект метрик для запроса
        """
        metrics = RequestMetrics(
            request_type=request_type,
            start_time=time.time()
        )
        
        # Отслеживаем память, если включено
        if self.memory_tracking_enabled:
            try:
                current, peak = tracemalloc.get_traced_memory()
                metrics.memory_peak_mb = peak / (1024 * 1024)  # Конвертируем в МБ
            except Exception:
                pass
        
        return metrics
    
    def end_request(
        self,
        metrics: RequestMetrics,
        success: bool = True,
        error_type: Optional[str] = None,
        cache_hit: bool = False,
        records_count: int = 0
    ) -> None:
        """
        Завершает отслеживание метрик для запроса.
        
        Args:
            metrics: Объект метрик
            success: Успешность запроса
            error_type: Тип ошибки (если была)
            cache_hit: Попадание в кеш
            records_count: Количество обработанных записей
        """
        metrics.end_time = time.time()
        metrics.duration_ms = (metrics.end_time - metrics.start_time) * 1000
        metrics.success = success
        metrics.error_type = error_type
        metrics.cache_hit = cache_hit
        metrics.records_count = records_count
        
        # Обновляем память
        if self.memory_tracking_enabled:
            try:
                current, peak = tracemalloc.get_traced_memory()
                metrics.memory_peak_mb = peak / (1024 * 1024)
            except Exception:
                pass
        
        # Сохраняем метрики
        self.metrics_history.append(metrics)
        
        # Обновляем счетчики
        self.request_counters[metrics.request_type] += 1
        self.total_duration_by_type[metrics.request_type] += metrics.duration_ms
        
        if not success and error_type:
            self.error_counters[error_type] += 1
        
        if cache_hit:
            self.cache_hit_counters[metrics.request_type] += 1
        
        # Ограничиваем историю (храним последние 1000 запросов)
        if len(self.metrics_history) > 1000:
            self.metrics_history = self.metrics_history[-1000:]
    
    def get_statistics(self, request_type: Optional[str] = None) -> Dict[str, Any]:
        """
        Получает статистику по метрикам.
        
        Args:
            request_type: Тип запроса для фильтрации (если None, по всем)
            
        Returns:
            Словарь со статистикой
        """
        if request_type:
            filtered_metrics = [m for m in self.metrics_history if m.request_type == request_type]
        else:
            filtered_metrics = self.metrics_history
        
        if not filtered_metrics:
            return {
                'total_requests': 0,
                'request_type': request_type
            }
        
        durations = [m.duration_ms for m in filtered_metrics if m.duration_ms is not None]
        memory_peaks = [m.memory_peak_mb for m in filtered_metrics if m.memory_peak_mb is not None]
        successful = [m for m in filtered_metrics if m.success]
        failed = [m for m in filtered_metrics if not m.success]
        cache_hits = [m for m in filtered_metrics if m.cache_hit]
        
        stats = {
            'total_requests': len(filtered_metrics),
            'request_type': request_type or 'all',
            'successful_requests': len(successful),
            'failed_requests': len(failed),
            'success_rate': (len(successful) / len(filtered_metrics) * 100) if filtered_metrics else 0,
            'cache_hits': len(cache_hits),
            'cache_hit_rate': (len(cache_hits) / len(filtered_metrics) * 100) if filtered_metrics else 0,
        }
        
        if durations:
            stats['avg_duration_ms'] = sum(durations) / len(durations)
            stats['min_duration_ms'] = min(durations)
            stats['max_duration_ms'] = max(durations)
            stats['median_duration_ms'] = sorted(durations)[len(durations) // 2]
        
        if memory_peaks:
            stats['avg_memory_mb'] = sum(memory_peaks) / len(memory_peaks)
            stats['max_memory_mb'] = max(memory_peaks)
        
        # Статистика по типам ошибок
        if failed:
            error_types = {}
            for m in failed:
                if m.error_type:
                    error_types[m.error_type] = error_types.get(m.error_type, 0) + 1
            stats['error_types'] = error_types
        
        # Статистика по количеству записей
        records_counts = [m.records_count for m in filtered_metrics if m.records_count > 0]
        if records_counts:
            stats['avg_records_per_request'] = sum(records_counts) / len(records_counts)
            stats['max_records_per_request'] = max(records_counts)
            stats['total_records_processed'] = sum(records_counts)
        
        return stats
    
    def get_summary(self) -> Dict[str, Any]:
        """
        Получает общую сводку по всем метрикам.
        
        Returns:
            Словарь со сводкой
        """
        summary = {
            'total_requests': len(self.metrics_history),
            'request_types': dict(self.request_counters),
            'error_types': dict(self.error_counters),
            'cache_hits_by_type': dict(self.cache_hit_counters),
            'avg_duration_by_type': {},
            'timestamp': datetime.now().isoformat()
        }
        
        # Средняя длительность по типам
        for request_type, total_duration in self.total_duration_by_type.items():
            count = self.request_counters.get(request_type, 1)
            summary['avg_duration_by_type'][request_type] = total_duration / count
        
        # Общая статистика
        if self.metrics_history:
            all_durations = [m.duration_ms for m in self.metrics_history if m.duration_ms is not None]
            if all_durations:
                summary['overall_avg_duration_ms'] = sum(all_durations) / len(all_durations)
                summary['overall_min_duration_ms'] = min(all_durations)
                summary['overall_max_duration_ms'] = max(all_durations)
            
            successful = len([m for m in self.metrics_history if m.success])
            summary['overall_success_rate'] = (successful / len(self.metrics_history)) * 100
            
            cache_hits = len([m for m in self.metrics_history if m.cache_hit])
            summary['overall_cache_hit_rate'] = (cache_hits / len(self.metrics_history)) * 100 if self.metrics_history else 0
        
        return summary
    
    def reset(self) -> None:
        """Сбрасывает все метрики."""
        self.metrics_history.clear()
        self.request_counters.clear()
        self.error_counters.clear()
        self.cache_hit_counters.clear()
        self.total_duration_by_type.clear()
        logger.info("Метрики сброшены")
    
    def get_recent_metrics(self, limit: int = 10) -> List[Dict[str, Any]]:
        """
        Получает последние метрики.
        
        Args:
            limit: Количество последних метрик
            
        Returns:
            Список словарей с метриками
        """
        recent = self.metrics_history[-limit:]
        return [
            {
                'request_type': m.request_type,
                'duration_ms': m.duration_ms,
                'success': m.success,
                'cache_hit': m.cache_hit,
                'records_count': m.records_count,
                'memory_peak_mb': m.memory_peak_mb
            }
            for m in recent
        ]


# Глобальный экземпляр сборщика метрик
_metrics_collector: Optional[MetricsCollector] = None


def get_metrics_collector() -> MetricsCollector:
    """
    Получает глобальный экземпляр сборщика метрик.
    
    Returns:
        Экземпляр MetricsCollector
    """
    global _metrics_collector
    if _metrics_collector is None:
        _metrics_collector = MetricsCollector()
    return _metrics_collector


def reset_metrics() -> None:
    """Сбрасывает все метрики."""
    collector = get_metrics_collector()
    collector.reset()

