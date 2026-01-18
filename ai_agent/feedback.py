"""
Модуль для сбора обратной связи от пользователей.
Позволяет пользователям оценивать качество ответов агента.
"""

import logging
from dataclasses import dataclass
from datetime import datetime
from typing import Any, Dict, List, Optional

logger = logging.getLogger(__name__)


@dataclass
class Feedback:
    """Обратная связь от пользователя."""
    user_id: Optional[int]
    message: str
    response: str
    rating: int  # 1-5
    comment: Optional[str] = None
    timestamp: datetime = None
    request_type: str = 'question'  # 'logistics', 'materials', 'duty', 'analytics', 'question'
    
    def __post_init__(self):
        if self.timestamp is None:
            self.timestamp = datetime.now()


class FeedbackCollector:
    """Класс для сбора обратной связи от пользователей."""
    
    def __init__(self):
        """Инициализирует сборщик обратной связи."""
        self.feedback_history: List[Feedback] = []
        self.ratings_by_type: Dict[str, List[int]] = {}
        self.total_feedback_count = 0
        self.average_rating = 0.0
    
    def add_feedback(
        self,
        user_id: Optional[int],
        message: str,
        response: str,
        rating: int,
        comment: Optional[str] = None,
        request_type: str = 'question'
    ) -> bool:
        """
        Добавляет обратную связь от пользователя.
        
        Args:
            user_id: ID пользователя
            message: Сообщение пользователя
            response: Ответ агента
            rating: Оценка (1-5)
            comment: Комментарий (опционально)
            request_type: Тип запроса
            
        Returns:
            True если обратная связь добавлена успешно
        """
        try:
            # Валидация оценки
            if not (1 <= rating <= 5):
                logger.warning(f"Некорректная оценка: {rating}, должна быть от 1 до 5")
                return False
            
            feedback = Feedback(
                user_id=user_id,
                message=message,
                response=response,
                rating=rating,
                comment=comment,
                request_type=request_type
            )
            
            self.feedback_history.append(feedback)
            
            # Обновляем статистику по типам
            if request_type not in self.ratings_by_type:
                self.ratings_by_type[request_type] = []
            self.ratings_by_type[request_type].append(rating)
            
            # Обновляем общую статистику
            self.total_feedback_count += 1
            all_ratings = [f.rating for f in self.feedback_history]
            self.average_rating = sum(all_ratings) / len(all_ratings) if all_ratings else 0.0
            
            logger.info(f"Добавлена обратная связь: оценка {rating}, тип {request_type}")
            return True
        
        except Exception as e:
            logger.error(f"Ошибка при добавлении обратной связи: {e}", exc_info=True)
            return False
    
    def get_statistics(self, request_type: Optional[str] = None) -> Dict[str, Any]:
        """
        Получает статистику по обратной связи.
        
        Args:
            request_type: Тип запроса для фильтрации (если None, по всем)
            
        Returns:
            Словарь со статистикой
        """
        if request_type:
            filtered_feedback = [f for f in self.feedback_history if f.request_type == request_type]
        else:
            filtered_feedback = self.feedback_history
        
        if not filtered_feedback:
            return {
                'total_count': 0,
                'request_type': request_type or 'all'
            }
        
        ratings = [f.rating for f in filtered_feedback]
        
        stats = {
            'total_count': len(filtered_feedback),
            'request_type': request_type or 'all',
            'average_rating': sum(ratings) / len(ratings),
            'min_rating': min(ratings),
            'max_rating': max(ratings),
            'rating_distribution': {
                1: len([r for r in ratings if r == 1]),
                2: len([r for r in ratings if r == 2]),
                3: len([r for r in ratings if r == 3]),
                4: len([r for r in ratings if r == 4]),
                5: len([r for r in ratings if r == 5])
            }
        }
        
        # Количество комментариев
        comments_count = len([f for f in filtered_feedback if f.comment])
        stats['comments_count'] = comments_count
        stats['comments_percentage'] = (comments_count / len(filtered_feedback) * 100) if filtered_feedback else 0
        
        return stats
    
    def get_recent_feedback(self, limit: int = 10) -> List[Dict[str, Any]]:
        """
        Получает последние отзывы.
        
        Args:
            limit: Количество последних отзывов
            
        Returns:
            Список словарей с обратной связью
        """
        recent = self.feedback_history[-limit:]
        return [
            {
                'user_id': f.user_id,
                'rating': f.rating,
                'comment': f.comment,
                'request_type': f.request_type,
                'timestamp': f.timestamp.isoformat() if f.timestamp else None,
                'message_preview': f.message[:50] + '...' if len(f.message) > 50 else f.message
            }
            for f in recent
        ]
    
    def get_summary(self) -> Dict[str, Any]:
        """
        Получает общую сводку по обратной связи.
        
        Returns:
            Словарь со сводкой
        """
        summary = {
            'total_feedback_count': self.total_feedback_count,
            'average_rating': self.average_rating,
            'ratings_by_type': {},
            'overall_rating_distribution': {
                1: 0, 2: 0, 3: 0, 4: 0, 5: 0
            },
            'timestamp': datetime.now().isoformat()
        }
        
        # Статистика по типам
        for request_type, ratings in self.ratings_by_type.items():
            if ratings:
                summary['ratings_by_type'][request_type] = {
                    'count': len(ratings),
                    'average': sum(ratings) / len(ratings)
                }
        
        # Общее распределение оценок
        for feedback in self.feedback_history:
            summary['overall_rating_distribution'][feedback.rating] += 1
        
        return summary
    
    def reset(self) -> None:
        """Сбрасывает всю обратную связь."""
        self.feedback_history.clear()
        self.ratings_by_type.clear()
        self.total_feedback_count = 0
        self.average_rating = 0.0
        logger.info("Обратная связь сброшена")


# Глобальный экземпляр сборщика обратной связи
_feedback_collector: Optional[FeedbackCollector] = None


def get_feedback_collector() -> FeedbackCollector:
    """
    Получает глобальный экземпляр сборщика обратной связи.
    
    Returns:
        Экземпляр FeedbackCollector
    """
    global _feedback_collector
    if _feedback_collector is None:
        _feedback_collector = FeedbackCollector()
    return _feedback_collector


def reset_feedback() -> None:
    """Сбрасывает всю обратную связь."""
    collector = get_feedback_collector()
    collector.reset()

