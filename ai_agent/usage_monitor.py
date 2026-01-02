"""
Мониторинг использования AI агента и OpenRouter API.
"""

import logging
from datetime import datetime, timedelta
from typing import Dict, List, Optional

from sqlalchemy import func, case
from sqlalchemy.exc import SQLAlchemyError

from app.core.extensions import SessionLocal
from app.models.models import AIAgentUsageRecord

logger = logging.getLogger(__name__)


class UsageStats:
    """Статистика использования AI агента."""
    
    def __init__(self, data: Dict):
        self.total_requests = data.get('total_requests', 0)
        self.total_tokens = data.get('total_tokens', 0)
        self.total_cost = data.get('total_cost', 0.0)
        self.avg_response_time = data.get('avg_response_time', 0)
        self.errors_count = data.get('errors_count', 0)
        self.cache_hits = data.get('cache_hits', 0)
        self.cache_hit_rate = data.get('cache_hit_rate', 0.0)
    
    def to_dict(self) -> Dict:
        """Преобразует в словарь."""
        return {
            'total_requests': self.total_requests,
            'total_tokens': self.total_tokens,
            'total_cost': self.total_cost,
            'avg_response_time': self.avg_response_time,
            'errors_count': self.errors_count,
            'cache_hits': self.cache_hits,
            'cache_hit_rate': self.cache_hit_rate,
        }


class UsageMonitor:
    """Мониторинг использования AI агента."""
    
    @staticmethod
    def log_usage(
        user_id: Optional[int] = None,
        request_tokens: Optional[int] = None,
        response_tokens: Optional[int] = None,
        total_cost: Optional[float] = None,
        response_time_ms: Optional[int] = None,
        model_name: Optional[str] = None,
        error_type: Optional[str] = None,
        user_message: Optional[str] = None,
        cache_hit: bool = False
    ) -> bool:
        """
        Записывает информацию об использовании AI агента в БД.
        
        Args:
            user_id: ID пользователя (если авторизован)
            request_tokens: Количество токенов в запросе
            response_tokens: Количество токенов в ответе
            total_cost: Стоимость запроса
            response_time_ms: Время ответа в миллисекундах
            model_name: Название модели
            error_type: Тип ошибки (если была)
            user_message: Сообщение пользователя (для анализа)
            cache_hit: Был ли использован кеш
            
        Returns:
            True если запись успешна, False иначе
        """
        try:
            session = SessionLocal()
            # Проверяем, что сессия имеет привязку к базе данных
            if session.bind is None:
                logger.warning("SessionLocal не имеет привязки к базе данных. Пропускаем запись статистики.")
                return False
            
            usage_record = AIAgentUsageRecord(
                user_id=user_id,
                request_tokens=request_tokens,
                response_tokens=response_tokens,
                total_cost=total_cost,
                response_time_ms=response_time_ms,
                model_name=model_name,
                error_type=error_type,
                user_message=user_message[:500] if user_message else None,  # Ограничиваем длину
                cache_hit=cache_hit
            )
            
            session.add(usage_record)
            session.commit()
            logger.debug(f"Записано использование AI агента: tokens={request_tokens}+{response_tokens}, "
                        f"cost={total_cost}, time={response_time_ms}ms, cache_hit={cache_hit}")
            return True
            
        except SQLAlchemyError as e:
            logger.error(f"Ошибка записи статистики использования AI: {e}")
            if 'session' in locals():
                session.rollback()
            return False
        except Exception as e:
            logger.error(f"Неожиданная ошибка записи статистики использования AI: {e}")
            return False
        finally:
            if 'session' in locals():
                session.close()
                SessionLocal.remove()
    
    @staticmethod
    def get_stats(
        days: int = 30,
        user_id: Optional[int] = None
    ) -> UsageStats:
        """
        Получает статистику использования за период.
        
        Args:
            days: Количество дней для анализа
            user_id: ID пользователя (если нужна статистика конкретного пользователя)
            
        Returns:
            UsageStats с агрегированной статистикой
        """
        try:
            session = SessionLocal()
            # Проверяем, что сессия имеет привязку к базе данных
            if session.bind is None:
                logger.warning("SessionLocal не имеет привязки к базе данных. Возвращаем пустую статистику.")
                return UsageStats({})
            
            # Определяем период
            start_date = datetime.now() - timedelta(days=days)
            
            # Базовый запрос
            query = session.query(
                func.count(AIAgentUsageRecord.id).label('total_requests'),
                func.sum(AIAgentUsageRecord.request_tokens).label('request_tokens'),
                func.sum(AIAgentUsageRecord.response_tokens).label('response_tokens'),
                func.sum(AIAgentUsageRecord.total_cost).label('total_cost'),
                func.avg(AIAgentUsageRecord.response_time_ms).label('avg_response_time'),
                func.sum(case((AIAgentUsageRecord.error_type.isnot(None), 1), else_=0)).label('errors_count'),
                func.sum(case((AIAgentUsageRecord.cache_hit == True, 1), else_=0)).label('cache_hits')
            ).filter(
                AIAgentUsageRecord.timestamp >= start_date
            )
            
            # Фильтруем по пользователю если указан
            if user_id:
                query = query.filter(AIAgentUsageRecord.user_id == user_id)
            
            result = query.one()
            
            # Вычисляем агрегированные метрики
            total_requests = result.total_requests or 0
            request_tokens = result.request_tokens or 0
            response_tokens = result.response_tokens or 0
            total_tokens = request_tokens + response_tokens
            cache_hits = result.cache_hits or 0
            cache_hit_rate = (cache_hits / total_requests * 100) if total_requests > 0 else 0.0
            
            stats_data = {
                'total_requests': total_requests,
                'total_tokens': total_tokens,
                'total_cost': float(result.total_cost or 0.0),
                'avg_response_time': int(result.avg_response_time or 0),
                'errors_count': result.errors_count or 0,
                'cache_hits': cache_hits,
                'cache_hit_rate': round(cache_hit_rate, 2)
            }
            
            return UsageStats(stats_data)
            
        except SQLAlchemyError as e:
            logger.error(f"Ошибка получения статистики использования AI: {e}")
            return UsageStats({})
        except Exception as e:
            logger.error(f"Неожиданная ошибка получения статистики использования AI: {e}")
            return UsageStats({})
        finally:
            if 'session' in locals():
                session.close()
                SessionLocal.remove()
    
    @staticmethod
    def get_hourly_stats(hours: int = 24) -> List[Dict]:
        """
        Получает почасовую статистику запросов.
        
        Args:
            hours: Количество часов для анализа
            
        Returns:
            Список словарей с почасовой статистикой
        """
        try:
            session = SessionLocal()
            # Проверяем, что сессия имеет привязку к базе данных
            if session.bind is None:
                logger.warning("SessionLocal не имеет привязки к базе данных. Возвращаем пустую статистику.")
                return []
            
            start_date = datetime.now() - timedelta(hours=hours)
            
            # Получаем все записи за период
            records = session.query(
                AIAgentUsageRecord.timestamp,
                AIAgentUsageRecord.error_type
            ).filter(
                AIAgentUsageRecord.timestamp >= start_date
            ).all()
            
            # Группируем по часам в Python (совместимо с SQLite)
            hourly_data = {}
            for record in records:
                # Форматируем timestamp до часа
                hour_key = record.timestamp.strftime('%Y-%m-%d %H:00:00')
                if hour_key not in hourly_data:
                    hourly_data[hour_key] = {'requests': 0, 'errors': 0}
                hourly_data[hour_key]['requests'] += 1
                if record.error_type is not None:
                    hourly_data[hour_key]['errors'] += 1
            
            # Преобразуем в список и сортируем
            results = [
                {
                    'hour': hour,
                    'requests': data['requests'],
                    'errors': data['errors']
                }
                for hour, data in sorted(hourly_data.items())
            ]
            
            return results
            
        except SQLAlchemyError as e:
            logger.error(f"Ошибка получения почасовой статистики AI: {e}")
            return []
        except Exception as e:
            logger.error(f"Неожиданная ошибка получения почасовой статистики AI: {e}")
            return []
        finally:
            if 'session' in locals():
                session.close()
                SessionLocal.remove()
    
    @staticmethod
    def get_error_stats(days: int = 7) -> List[Dict]:
        """
        Получает статистику ошибок по типам.
        
        Args:
            days: Количество дней для анализа
            
        Returns:
            Список словарей с типами ошибок и их количеством
        """
        try:
            session = SessionLocal()
            # Проверяем, что сессия имеет привязку к базе данных
            if session.bind is None:
                logger.warning("SessionLocal не имеет привязки к базе данных. Возвращаем пустую статистику.")
                return []
            
            start_date = datetime.now() - timedelta(days=days)
            
            query = session.query(
                AIAgentUsageRecord.error_type,
                func.count(AIAgentUsageRecord.id).label('count')
            ).filter(
                AIAgentUsageRecord.timestamp >= start_date,
                AIAgentUsageRecord.error_type.isnot(None)
            ).group_by(
                AIAgentUsageRecord.error_type
            ).order_by(func.count(AIAgentUsageRecord.id).desc())
            
            results = query.all()
            
            return [
                {
                    'error_type': row.error_type,
                    'count': row.count
                }
                for row in results
            ]
            
        except SQLAlchemyError as e:
            logger.error(f"Ошибка получения статистики ошибок AI: {e}")
            return []
        except Exception as e:
            logger.error(f"Неожиданная ошибка получения статистики ошибок AI: {e}")
            return []
        finally:
            if 'session' in locals():
                session.close()
                SessionLocal.remove()
    
    @staticmethod
    def get_top_queries(limit: int = 10, days: int = 7) -> List[Dict]:
        """
        Получает топ самых частых запросов.
        
        Args:
            limit: Количество запросов в топе
            days: Количество дней для анализа
            
        Returns:
            Список словарей с запросами и их частотой
        """
        try:
            session = SessionLocal()
            # Проверяем, что сессия имеет привязку к базе данных
            if session.bind is None:
                logger.warning("SessionLocal не имеет привязки к базе данных. Возвращаем пустую статистику.")
                return []
            
            start_date = datetime.now() - timedelta(days=days)
            
            query = session.query(
                AIAgentUsageRecord.user_message,
                func.count(AIAgentUsageRecord.id).label('count')
            ).filter(
                AIAgentUsageRecord.timestamp >= start_date,
                AIAgentUsageRecord.user_message.isnot(None)
            ).group_by(
                AIAgentUsageRecord.user_message
            ).order_by(
                func.count(AIAgentUsageRecord.id).desc()
            ).limit(limit)
            
            results = query.all()
            
            return [
                {
                    'query': row.user_message[:100],  # Ограничиваем длину для отображения
                    'count': row.count
                }
                for row in results
            ]
            
        except SQLAlchemyError as e:
            logger.error(f"Ошибка получения топа запросов AI: {e}")
            return []
        except Exception as e:
            logger.error(f"Неожиданная ошибка получения топа запросов AI: {e}")
            return []
        finally:
            if 'session' in locals():
                session.close()
                SessionLocal.remove()
    
    @staticmethod
    def get_model_usage_stats(days: int = 30) -> List[Dict]:
        """
        Получает статистику использования по моделям.
        
        Args:
            days: Количество дней для анализа
            
        Returns:
            Список словарей со статистикой по моделям
        """
        try:
            session = SessionLocal()
            # Проверяем, что сессия имеет привязку к базе данных
            if session.bind is None:
                logger.warning("SessionLocal не имеет привязки к базе данных. Возвращаем пустую статистику.")
                return []
            
            start_date = datetime.now() - timedelta(days=days)
            
            query = session.query(
                AIAgentUsageRecord.model_name,
                func.count(AIAgentUsageRecord.id).label('requests'),
                func.sum(AIAgentUsageRecord.total_cost).label('total_cost'),
                func.avg(AIAgentUsageRecord.response_time_ms).label('avg_time')
            ).filter(
                AIAgentUsageRecord.timestamp >= start_date,
                AIAgentUsageRecord.model_name.isnot(None)
            ).group_by(
                AIAgentUsageRecord.model_name
            ).order_by(func.count(AIAgentUsageRecord.id).desc())
            
            results = query.all()
            
            return [
                {
                    'model_name': row.model_name,
                    'requests': row.requests,
                    'total_cost': float(row.total_cost or 0.0),
                    'avg_time_ms': int(row.avg_time or 0)
                }
                for row in results
            ]
            
        except SQLAlchemyError as e:
            logger.error(f"Ошибка получения статистики по моделям: {e}")
            return []
        except Exception as e:
            logger.error(f"Неожиданная ошибка получения статистики по моделям: {e}")
            return []
        finally:
            if 'session' in locals():
                session.close()
                SessionLocal.remove()


# Удобная функция для быстрого логирования
def log_ai_usage(**kwargs) -> bool:
    """
    Удобная функция для логирования использования AI агента.
    
    Args:
        **kwargs: Параметры для передачи в UsageMonitor.log_usage()
        
    Returns:
        True если запись успешна
    """
    return UsageMonitor.log_usage(**kwargs)


def get_api_key_status_from_db() -> Optional[str]:
    """
    Получает последнюю ошибку API ключа из базы данных.
    
    Returns:
        Строка с описанием последней ошибки или None, если ошибок нет
    """
    try:
        session = SessionLocal()
        # Проверяем, что сессия имеет привязку к базе данных
        if session.bind is None:
            logger.warning("SessionLocal не имеет привязки к базе данных. Не можем получить статус API ключа.")
            return None
        
        # Получаем последнюю запись с ошибкой
        last_error = session.query(
            AIAgentUsageRecord.error_type,
            AIAgentUsageRecord.timestamp
        ).filter(
            AIAgentUsageRecord.error_type.isnot(None)
        ).order_by(
            AIAgentUsageRecord.timestamp.desc()
        ).first()
        
        if last_error:
            # Форматируем сообщение об ошибке
            error_type = last_error.error_type
            timestamp = last_error.timestamp.strftime('%Y-%m-%d %H:%M:%S') if last_error.timestamp else 'неизвестно'
            return f"{error_type} (время: {timestamp})"
        
        return None
        
    except SQLAlchemyError as e:
        logger.error(f"Ошибка получения статуса API ключа из БД: {e}")
        return None
    except Exception as e:
        logger.error(f"Неожиданная ошибка получения статуса API ключа: {e}")
        return None
    finally:
        if 'session' in locals():
            session.close()
            SessionLocal.remove()

