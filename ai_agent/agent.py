"""
Основной модуль AI агента с интеграцией OpenRouter API.
"""

import json
import re
import sys
from typing import Any, Dict, List, Optional

import requests

from ai_agent.config import (
    get_api_key,
    get_api_url,
    get_model_name,
    is_reasoning_enabled,
)
from ai_agent.knowledge_base import KnowledgeBase
from ai_agent.logistics_helper import LogisticsHelper


class AIAgent:
    """AI агент для консультаций по проекту."""
    
    def __init__(self, api_key: Optional[str] = None, model: Optional[str] = None):
        """
        Инициализирует AI агента.
        
        Args:
            api_key: API ключ OpenRouter (если None, берется из config)
            model: Название модели (если None, берется из config)
        """
        self.api_key = api_key or get_api_key()
        self.api_url = get_api_url()
        self.model = model or get_model_name()
        self.reasoning_enabled = is_reasoning_enabled()
        
        # Инициализируем помощников
        self.knowledge_base = KnowledgeBase()
        self.logistics_helper = LogisticsHelper()
        
        # История диалога
        self.conversation_history: List[Dict[str, Any]] = []
        
        # Системный промпт
        self.system_prompt = self._build_system_prompt()
    
    def _build_system_prompt(self) -> str:
        """
        Формирует системный промпт для агента.
        
        Returns:
            Системный промпт
        """
        # Получаем краткий контекст документации
        docs_context = self.knowledge_base.get_all_context()
        
        prompt = """Ты - AI консультант и помощник для проекта KP Generator (генератор коммерческих предложений).

Твоя роль:
- Отвечать на вопросы по документации проекта (инструкции, руководства)
- Помогать с расчетом логистики
- Консультировать по использованию системы

Доступные функции:
1. Ответы на вопросы по документации - используй контекст из базы знаний
2. Расчет логистики - можешь вызывать функцию расчета для грузов

Правила работы с логистикой:
- Если пользователь просит рассчитать логистику, извлекай параметры:
  * Вес груза (в кг)
  * Город назначения
  * Тип транспорта (фура/трал, по умолчанию фура)
  * Габариты (опционально: длина, ширина, высота в мм)
- Если параметров недостаточно, уточни у пользователя
- После расчета логистики, представь результат в понятном формате

Правила ответов:
- Отвечай на русском языке
- Будь кратким, но информативным
- Используй контекст из документации для точных ответов
- Если не знаешь ответа, честно скажи об этом

Контекст документации:
"""
        prompt += docs_context[:5000]  # Ограничиваем размер
        
        return prompt
    
    def _extract_intent(self, message: str) -> Dict[str, Any]:
        """
        Определяет намерение пользователя из сообщения.
        
        Args:
            message: Сообщение пользователя
            
        Returns:
            Словарь с информацией о намерении
        """
        message_lower = message.lower()
        
        # Ключевые слова для расчета логистики
        logistics_keywords = [
            'логистик', 'рассчитай', 'расчет', 'доставк', 'перевозк',
            'стоимость доставки', 'цена доставки', 'логистика для',
            'сколько стоит', 'стоимость перевозки'
        ]
        
        is_logistics = any(keyword in message_lower for keyword in logistics_keywords)
        
        # Попытка извлечь параметры логистики
        logistics_params = {}
        if is_logistics:
            # Вес
            weight_match = re.search(r'(\d+(?:[.,]\d+)?)\s*(?:кг|kg|тонн|т)', message_lower)
            if weight_match:
                weight_str = weight_match.group(1).replace(',', '.')
                try:
                    weight = float(weight_str)
                    # Если указано в тоннах, переводим в кг
                    if 'тонн' in message_lower or 'т' in message_lower:
                        weight *= 1000
                    logistics_params['weight_kg'] = weight
                except ValueError:
                    pass
            
            # Город (поиск по словам, включая сокращения)
            cities = [c.get('name') for c in self.logistics_helper._load_cities()]
            # Словарь сокращений
            city_aliases = {
                'екб': 'Екатеринбург',
                'мск': 'Москва',
                'спб': 'Санкт-Петербург',
                'питер': 'Санкт-Петербург',
            }
            
            # Сначала проверяем сокращения
            for alias, city_name in city_aliases.items():
                if alias in message_lower and city_name in cities:
                    logistics_params['city_name'] = city_name
                    break
            
            # Если не нашли по сокращению, ищем по полному названию
            if 'city_name' not in logistics_params:
                for city in cities:
                    city_lower = city.lower()
                    # Проверяем точное вхождение или начало слова
                    if city_lower in message_lower:
                        logistics_params['city_name'] = city
                        break
            
            # Тип транспорта
            if 'трал' in message_lower:
                logistics_params['transport_type'] = 'trail'
            else:
                logistics_params['transport_type'] = 'truck'
            
            # Габариты (опционально)
            dim_match = re.search(r'(\d+)\s*[xх×]\s*(\d+)\s*[xх×]\s*(\d+)\s*(?:мм|mm)', message_lower)
            if dim_match:
                logistics_params['length_mm'] = float(dim_match.group(1))
                logistics_params['width_mm'] = float(dim_match.group(2))
                logistics_params['height_mm'] = float(dim_match.group(3))
        
        return {
            'is_logistics': is_logistics,
            'logistics_params': logistics_params,
            'is_question': '?' in message or any(word in message_lower for word in ['как', 'что', 'почему', 'где', 'когда', 'зачем']),
        }
    
    def _call_api(self, messages: List[Dict[str, Any]]) -> Dict[str, Any]:
        """
        Вызывает API OpenRouter.
        
        Args:
            messages: Список сообщений для отправки
            
        Returns:
            Ответ от API
        """
        headers = {
            "Authorization": f"Bearer {self.api_key}",
            "Content-Type": "application/json",
        }
        
        payload = {
            "model": self.model,
            "messages": messages,
        }
        
        if self.reasoning_enabled:
            payload["reasoning"] = {"enabled": True}
        
        try:
            response = requests.post(
                url=self.api_url,
                headers=headers,
                data=json.dumps(payload),
                timeout=60
            )
            response.raise_for_status()
            return response.json()
        except requests.exceptions.RequestException as e:
            raise Exception(f"Ошибка API: {str(e)}")
    
    def chat(self, message: str, context: Optional[Dict[str, Any]] = None) -> str:
        """
        Основной метод для общения с агентом.
        
        Args:
            message: Сообщение пользователя
            context: Дополнительный контекст (не используется пока)
            
        Returns:
            Ответ агента
        """
        # Определяем намерение
        intent = self._extract_intent(message)
        
        # Если это запрос на расчет логистики и есть параметры
        if intent['is_logistics'] and intent['logistics_params'].get('weight_kg'):
            params = intent['logistics_params']
            result = self.logistics_helper.calculate_simple_logistics(
                weight_kg=params.get('weight_kg', 0),
                city_name=params.get('city_name', ''),
                transport_type=params.get('transport_type', 'truck'),
                length_mm=params.get('length_mm'),
                width_mm=params.get('width_mm'),
                height_mm=params.get('height_mm'),
            )
            
            # Форматируем результат
            formatted_result = self.logistics_helper.format_logistics_result(result)
            
            # Если есть ошибка или недостаточно параметров, возвращаем сразу
            if 'error' in result or not params.get('city_name'):
                # Добавляем в историю для контекста
                self.conversation_history.append({
                    "role": "user",
                    "content": message
                })
                self.conversation_history.append({
                    "role": "assistant",
                    "content": formatted_result
                })
                return formatted_result
            
            # Для успешного расчета возвращаем отформатированный результат напрямую
            # (он уже содержит все необходимые данные)
            # Добавляем в историю для контекста
            self.conversation_history.append({
                "role": "user",
                "content": message
            })
            self.conversation_history.append({
                "role": "assistant",
                "content": formatted_result
            })
            
            return formatted_result
        
        # Обычный вопрос - используем AI с контекстом
        # Добавляем в историю
        self.conversation_history.append({
            "role": "user",
            "content": message
        })
        
        # Получаем релевантный контекст
        relevant_context = self.knowledge_base.get_relevant_context(message)
        
        # Формируем сообщения для API
        api_messages = [
            {"role": "system", "content": self.system_prompt},
        ]
        
        # Добавляем релевантный контекст, если есть
        if relevant_context:
            user_content = f"Контекст из документации:\n{relevant_context}\n\nВопрос пользователя: {message}"
        else:
            user_content = message
        
        api_messages.append({
            "role": "user",
            "content": user_content
        })
        
        # Добавляем историю (последние 5 сообщений)
        # Форматируем для API (reasoning_details передаем только для assistant)
        for hist_msg in self.conversation_history[-5:]:
            msg_for_api = {
                "role": hist_msg["role"],
                "content": hist_msg["content"]
            }
            # Добавляем reasoning_details только для assistant сообщений
            if hist_msg["role"] == "assistant" and "reasoning_details" in hist_msg:
                msg_for_api["reasoning_details"] = hist_msg["reasoning_details"]
            api_messages.append(msg_for_api)
        
        try:
            response_data = self._call_api(api_messages)
            assistant_message = response_data['choices'][0]['message']
            content = assistant_message.get('content', '')
            
            # Сохраняем ответ в историю
            assistant_msg = {
                "role": "assistant",
                "content": content
            }
            
            # Сохраняем reasoning_details если есть (для следующего запроса)
            if 'reasoning_details' in assistant_message:
                assistant_msg["reasoning_details"] = assistant_message.get('reasoning_details')
            
            self.conversation_history.append(assistant_msg)
            
            return content
        except Exception as e:
            return f"Ошибка при обращении к AI: {str(e)}"
    
    def clear_history(self) -> None:
        """Очищает историю диалога."""
        self.conversation_history = []

