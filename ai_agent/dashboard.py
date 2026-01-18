"""
Модуль для создания dashboard с метриками использования агента.
Включает графики использования, статистику кеша, популярные запросы.
"""

import logging
from datetime import datetime, timedelta
from typing import Any, Dict, List, Optional

logger = logging.getLogger(__name__)


class DashboardGenerator:
    """Класс для генерации dashboard с метриками."""
    
    def __init__(self, metrics_collector=None, feedback_collector=None):
        """
        Инициализирует генератор dashboard.
        
        Args:
            metrics_collector: Экземпляр MetricsCollector
            feedback_collector: Экземпляр FeedbackCollector
        """
        self.metrics_collector = metrics_collector
        self.feedback_collector = feedback_collector
    
    def generate_dashboard_summary(
        self,
        days: int = 7
    ) -> Dict[str, Any]:
        """
        Генерирует общую сводку dashboard.
        
        Args:
            days: Количество дней для анализа
            
        Returns:
            Словарь со сводкой dashboard
        """
        summary = {
            'period_days': days,
            'timestamp': datetime.now().isoformat(),
            'metrics': {},
            'feedback': {},
            'cache': {},
            'popular_queries': []
        }
        
        # Метрики производительности
        if self.metrics_collector:
            metrics_summary = self.metrics_collector.get_summary()
            summary['metrics'] = {
                'total_requests': metrics_summary.get('total_requests', 0),
                'success_rate': metrics_summary.get('overall_success_rate', 0),
                'avg_duration_ms': metrics_summary.get('overall_avg_duration_ms', 0),
                'cache_hit_rate': metrics_summary.get('overall_cache_hit_rate', 0),
                'requests_by_type': metrics_summary.get('request_types', {}),
                'errors_by_type': metrics_summary.get('error_types', {})
            }
        
        # Обратная связь
        if self.feedback_collector:
            feedback_summary = self.feedback_collector.get_summary()
            summary['feedback'] = {
                'total_feedback': feedback_summary.get('total_feedback_count', 0),
                'average_rating': feedback_summary.get('average_rating', 0),
                'ratings_by_type': feedback_summary.get('ratings_by_type', {})
            }
        
        # Популярные запросы (из метрик)
        if self.metrics_collector:
            request_types = self.metrics_collector.request_counters
            if request_types:
                # Сортируем по частоте
                sorted_types = sorted(
                    request_types.items(),
                    key=lambda x: x[1],
                    reverse=True
                )
                summary['popular_queries'] = [
                    {'type': req_type, 'count': count}
                    for req_type, count in sorted_types[:10]
                ]
        
        return summary
    
    def generate_usage_chart_data(
        self,
        days: int = 7,
        group_by: str = 'day'  # 'day', 'hour'
    ) -> Dict[str, Any]:
        """
        Генерирует данные для графика использования.
        
        Args:
            days: Количество дней
            group_by: Группировка ('day' или 'hour')
            
        Returns:
            Словарь с данными для графика
        """
        if not self.metrics_collector:
            return {'error': 'Metrics collector not available'}
        
        # Получаем метрики за период
        cutoff_date = datetime.now() - timedelta(days=days)
        recent_metrics = [
            m for m in self.metrics_collector.metrics_history
            if m.start_time >= cutoff_date.timestamp()
        ]
        
        if not recent_metrics:
            return {'labels': [], 'data': []}
        
        # Группируем по периодам
        usage_by_period = {}
        for metric in recent_metrics:
            dt = datetime.fromtimestamp(metric.start_time)
            
            if group_by == 'day':
                period_key = dt.strftime('%Y-%m-%d')
            else:  # hour
                period_key = dt.strftime('%Y-%m-%d %H:00')
            
            if period_key not in usage_by_period:
                usage_by_period[period_key] = 0
            usage_by_period[period_key] += 1
        
        # Сортируем по дате
        sorted_periods = sorted(usage_by_period.items())
        
        return {
            'labels': [p[0] for p in sorted_periods],
            'data': [p[1] for p in sorted_periods],
            'total': len(recent_metrics)
        }
    
    def generate_performance_stats(
        self,
        request_type: Optional[str] = None
    ) -> Dict[str, Any]:
        """
        Генерирует статистику производительности.
        
        Args:
            request_type: Тип запроса для фильтрации (если None, по всем)
            
        Returns:
            Словарь со статистикой производительности
        """
        if not self.metrics_collector:
            return {'error': 'Metrics collector not available'}
        
        stats = self.metrics_collector.get_statistics(request_type)
        
        return {
            'total_requests': stats.get('total_requests', 0),
            'successful': stats.get('successful_requests', 0),
            'failed': stats.get('failed_requests', 0),
            'success_rate': stats.get('success_rate', 0),
            'cache_hits': stats.get('cache_hits', 0),
            'cache_hit_rate': stats.get('cache_hit_rate', 0),
            'avg_duration_ms': stats.get('avg_duration_ms', 0),
            'min_duration_ms': stats.get('min_duration_ms', 0),
            'max_duration_ms': stats.get('max_duration_ms', 0),
            'avg_memory_mb': stats.get('avg_memory_mb', 0),
            'max_memory_mb': stats.get('max_memory_mb', 0),
            'error_types': stats.get('error_types', {})
        }
    
    def generate_cache_stats(self) -> Dict[str, Any]:
        """
        Генерирует статистику кеша.
        
        Returns:
            Словарь со статистикой кеша
        """
        try:
            from ai_agent.cache_manager import AICacheManager
            cache_stats = AICacheManager.get_cache_stats()
            
            return {
                'enabled': cache_stats.get('enabled', False),
                'hit_rate': cache_stats.get('hit_rate', 0),
                'total_hits': cache_stats.get('total_hits', 0),
                'total_misses': cache_stats.get('total_misses', 0),
                'size_mb': cache_stats.get('size_mb', 0)
            }
        except Exception as e:
            logger.error(f"Ошибка при получении статистики кеша: {e}")
            return {'error': str(e)}
    
    def generate_feedback_stats(self) -> Dict[str, Any]:
        """
        Генерирует статистику обратной связи.
        
        Returns:
            Словарь со статистикой обратной связи
        """
        if not self.feedback_collector:
            return {'error': 'Feedback collector not available'}
        
        feedback_summary = self.feedback_collector.get_summary()
        
        return {
            'total_feedback': feedback_summary.get('total_feedback_count', 0),
            'average_rating': feedback_summary.get('average_rating', 0),
            'rating_distribution': feedback_summary.get('overall_rating_distribution', {}),
            'ratings_by_type': feedback_summary.get('ratings_by_type', {})
        }
    
    def format_dashboard_markdown(
        self,
        summary: Dict[str, Any]
    ) -> str:
        """
        Форматирует dashboard в markdown для отображения.
        
        Args:
            summary: Сводка dashboard
            
        Returns:
            Markdown строка
        """
        lines = ["# 📊 Dashboard метрик агента\n"]
        
        # Общая статистика
        metrics = summary.get('metrics', {})
        if metrics:
            lines.append("## Общая статистика\n")
            lines.append(f"**Всего запросов**: {metrics.get('total_requests', 0)}")
            lines.append(f"**Успешность**: {metrics.get('success_rate', 0):.1f}%")
            lines.append(f"**Среднее время ответа**: {metrics.get('avg_duration_ms', 0):.1f} мс")
            lines.append(f"**Попаданий в кеш**: {metrics.get('cache_hit_rate', 0):.1f}%")
            lines.append("")
        
        # Запросы по типам
        requests_by_type = metrics.get('requests_by_type', {})
        if requests_by_type:
            lines.append("## Запросы по типам\n")
            for req_type, count in sorted(requests_by_type.items(), key=lambda x: x[1], reverse=True):
                lines.append(f"- **{req_type}**: {count}")
            lines.append("")
        
        # Обратная связь
        feedback = summary.get('feedback', {})
        if feedback.get('total_feedback', 0) > 0:
            lines.append("## Обратная связь\n")
            lines.append(f"**Всего отзывов**: {feedback.get('total_feedback', 0)}")
            lines.append(f"**Средняя оценка**: {feedback.get('average_rating', 0):.1f}/5")
            lines.append("")
        
        # Популярные запросы
        popular_queries = summary.get('popular_queries', [])
        if popular_queries:
            lines.append("## Популярные запросы\n")
            for i, query_info in enumerate(popular_queries[:5], 1):
                lines.append(f"{i}. **{query_info['type']}**: {query_info['count']} запросов")
            lines.append("")
        
        lines.append(f"*Обновлено: {summary.get('timestamp', 'N/A')}*")
        
        return "\n".join(lines)

