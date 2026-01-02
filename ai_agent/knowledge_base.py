"""
Модуль для загрузки и управления базой знаний (документация и инструкции).
"""

import re
from pathlib import Path
from typing import Dict, List, Optional

from ai_agent.config import DOCS_DIR, INSTRUCTIONS_DIR


class KnowledgeBase:
    """Класс для управления базой знаний проекта."""
    
    def __init__(self):
        """Инициализирует базу знаний."""
        self.instructions: List[Dict[str, str]] = []
        self.documentation: List[Dict[str, str]] = []
        try:
            self._load_all()
        except Exception as e:
            import logging
            logging.error(f"Ошибка при загрузке базы знаний: {e}", exc_info=True)
            # Продолжаем работу с пустыми списками, чтобы не ломать инициализацию агента
            self.instructions = []
            self.documentation = []
    
    def _load_all(self) -> None:
        """Загружает всю документацию и инструкции."""
        try:
            self.instructions = self.load_instructions()
            self.documentation = self.load_documentation()
        except Exception as e:
            import logging
            logging.error(f"Ошибка в _load_all: {e}", exc_info=True)
            raise
    
    def load_instructions(self) -> List[Dict[str, str]]:
        """
        Загружает все инструкции из ai_agent/data/instructions/.
        
        Returns:
            Список словарей с ключами 'title' и 'content'
        """
        instructions = []
        
        if not INSTRUCTIONS_DIR.exists():
            import logging
            logging.warning(f'Instructions directory not found: {INSTRUCTIONS_DIR}')
            return instructions
        
        # Загружаем все .txt файлы из папки инструкций
        for file_path in sorted(INSTRUCTIONS_DIR.glob("*.txt")):
            try:
                with file_path.open('r', encoding='utf-8') as f:
                    content = f.read().strip()
                
                # Извлекаем название из первой строки или имени файла
                lines = content.split('\n')
                title = lines[0].strip() if lines else file_path.stem
                
                instructions.append({
                    'title': title,
                    'content': content,
                    'source': file_path.name
                })
            except Exception as e:
                import logging
                logging.error(f"Ошибка при загрузке инструкции {file_path}: {e}", exc_info=True)
        
        return instructions
    
    def load_documentation(self) -> List[Dict[str, str]]:
        """
        Загружает документацию из ai_agent/data/documentation/.
        
        Returns:
            Список словарей с ключами 'title' и 'content'
        """
        documentation = []
        
        if not DOCS_DIR.exists():
            import logging
            logging.warning(f'Documentation directory not found: {DOCS_DIR}')
            return documentation
        
        # Приоритетные файлы документации
        priority_files = [
            'COMPLETE_GUIDE.md',
            'SETUP.md',
            'PROJECT_STRUCTURE.md',
            'MULTI_POSITION_INTEGRATION.md',
            'USER_MANAGEMENT.md',
            'BUDGET_EXCEL_FLOW.md',
            'CHANGELOG.md',
            'LOGISTICS_CALCULATION.md',
            'MATERIALS_CATALOG.md',
            'DUTY_RATES.md',
        ]
        
        # Сначала загружаем приоритетные файлы
        for filename in priority_files:
            file_path = DOCS_DIR / filename
            if file_path.exists():
                doc = self._load_doc_file(file_path)
                if doc:
                    documentation.append(doc)
        
        # Затем загружаем остальные .md файлы
        for file_path in sorted(DOCS_DIR.glob("*.md")):
            if file_path.name not in priority_files:
                doc = self._load_doc_file(file_path)
                if doc:
                    documentation.append(doc)
        
        return documentation
    
    def _load_doc_file(self, file_path: Path) -> Optional[Dict[str, str]]:
        """
        Загружает один файл документации.
        
        Args:
            file_path: Путь к файлу
            
        Returns:
            Словарь с 'title' и 'content' или None при ошибке
        """
        try:
            with file_path.open('r', encoding='utf-8') as f:
                content = f.read().strip()
            
            # Извлекаем заголовок из первого # заголовка или используем имя файла
            title_match = re.search(r'^#\s+(.+)$', content, re.MULTILINE)
            title = title_match.group(1).strip() if title_match else file_path.stem
            
            return {
                'title': title,
                'content': content,
                'source': file_path.name
            }
        except Exception as e:
            import logging
            logging.error(f"Ошибка при загрузке документации {file_path}: {e}", exc_info=True)
            return None
    
    def get_relevant_context(self, query: str, max_results: int = 3) -> str:
        """
        Находит релевантные фрагменты документации по запросу.
        
        Args:
            query: Поисковый запрос
            max_results: Максимальное количество релевантных фрагментов
            
        Returns:
            Строка с релевантным контекстом
        """
        query_lower = query.lower()
        query_words = set(query_lower.split())
        
        # Ключевые фразы из запроса (для более точного поиска)
        key_phrases = []
        if 'согласование' in query_lower or 'согласовать' in query_lower:
            key_phrases.append('согласование')
        if 'бюджет' in query_lower:
            key_phrases.append('бюджет')
        if 'расценк' in query_lower or 'расценка' in query_lower:
            key_phrases.append('расценк')
        if 'отправ' in query_lower or 'отправить' in query_lower:
            key_phrases.append('отправ')
        
        # Собираем все документы
        all_docs = self.instructions + self.documentation
        
        # Улучшенный поиск по ключевым словам и фразам
        scored_docs = []
        for doc in all_docs:
            content_lower = doc['content'].lower()
            title_lower = doc['title'].lower()
            
            # Подсчитываем совпадения слов
            title_score = sum(1 for word in query_words if word in title_lower)
            content_score = sum(1 for word in query_words if word in content_lower)
            
            # Бонус за совпадение ключевых фраз
            phrase_bonus = 0
            for phrase in key_phrases:
                if phrase in content_lower:
                    phrase_bonus += 5  # Большой бонус за ключевые фразы
                if phrase in title_lower:
                    phrase_bonus += 10  # Еще больший бонус в заголовке
            
            # Общий score (заголовок важнее)
            total_score = title_score * 3 + content_score + phrase_bonus
            
            if total_score > 0:
                scored_docs.append((total_score, doc))
        
        # Сортируем по score
        scored_docs.sort(key=lambda x: x[0], reverse=True)
        
        # Формируем контекст
        context_parts = []
        for score, doc in scored_docs[:max_results]:
            content = doc['content']
            
            # Увеличиваем лимит до 3000 символов для инструкций
            max_chars = 3000 if doc['source'].endswith('.txt') else 2000
            
            if len(content) > max_chars:
                # Умное обрезание: ищем релевантные разделы
                lines = content.split('\n')
                relevant_lines = []
                char_count = 0
                found_relevant_section = False
                
                # Если есть ключевые фразы, ищем разделы с ними
                if key_phrases:
                    for i, line in enumerate(lines):
                        line_lower = line.lower()
                        # Проверяем, содержит ли строка ключевую фразу
                        if any(phrase in line_lower for phrase in key_phrases):
                            found_relevant_section = True
                            # Берем этот раздел и несколько строк до/после
                            start = max(0, i - 5)
                            end = min(len(lines), i + 30)  # Берем больше строк после найденного раздела
                            relevant_lines = lines[start:end]
                            break
                
                # Если не нашли релевантный раздел, берем начало
                if not relevant_lines:
                    for line in lines:
                        if char_count + len(line) > max_chars:
                            break
                        relevant_lines.append(line)
                        char_count += len(line) + 1
                    relevant_lines.append("...")
                
                content = '\n'.join(relevant_lines)
            
            context_parts.append(
                f"=== {doc['title']} ({doc['source']}) ===\n{content}\n"
            )
        
        return "\n".join(context_parts) if context_parts else ""
    
    def get_all_context(self) -> str:
        """
        Возвращает весь контекст документации для системного промпта.
        
        Returns:
            Полный контекст документации
        """
        parts = []
        
        if self.instructions:
            parts.append("=== ИНСТРУКЦИИ ===\n")
            for inst in self.instructions:
                parts.append(f"--- {inst['title']} ---\n{inst['content']}\n")
        
        if self.documentation:
            parts.append("\n=== ДОКУМЕНТАЦИЯ ===\n")
            for doc in self.documentation[:5]:  # Ограничиваем для системного промпта
                parts.append(f"--- {doc['title']} ---\n{doc['content'][:2000]}...\n")
        
        return "\n".join(parts)

