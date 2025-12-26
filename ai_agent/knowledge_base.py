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
        self._load_all()
    
    def _load_all(self) -> None:
        """Загружает всю документацию и инструкции."""
        self.instructions = self.load_instructions()
        self.documentation = self.load_documentation()
    
    def load_instructions(self) -> List[Dict[str, str]]:
        """
        Загружает все инструкции из static/instructions/.
        
        Returns:
            Список словарей с ключами 'title' и 'content'
        """
        instructions = []
        
        if not INSTRUCTIONS_DIR.exists():
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
                print(f"Ошибка при загрузке инструкции {file_path}: {e}")
        
        return instructions
    
    def load_documentation(self) -> List[Dict[str, str]]:
        """
        Загружает документацию из docs/.
        
        Returns:
            Список словарей с ключами 'title' и 'content'
        """
        documentation = []
        
        if not DOCS_DIR.exists():
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
            print(f"Ошибка при загрузке документации {file_path}: {e}")
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
        
        # Собираем все документы
        all_docs = self.instructions + self.documentation
        
        # Простой поиск по ключевым словам
        scored_docs = []
        for doc in all_docs:
            content_lower = doc['content'].lower()
            title_lower = doc['title'].lower()
            
            # Подсчитываем совпадения
            title_score = sum(1 for word in query_words if word in title_lower)
            content_score = sum(1 for word in query_words if word in content_lower)
            
            # Общий score (заголовок важнее)
            total_score = title_score * 3 + content_score
            
            if total_score > 0:
                scored_docs.append((total_score, doc))
        
        # Сортируем по score
        scored_docs.sort(key=lambda x: x[0], reverse=True)
        
        # Формируем контекст
        context_parts = []
        for score, doc in scored_docs[:max_results]:
            # Берем первые 1000 символов или до конца первого раздела
            content = doc['content']
            if len(content) > 1000:
                # Пытаемся обрезать по разделам
                lines = content.split('\n')
                truncated = []
                char_count = 0
                for line in lines:
                    if char_count + len(line) > 1000:
                        break
                    truncated.append(line)
                    char_count += len(line) + 1
                content = '\n'.join(truncated) + "..."
            
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

