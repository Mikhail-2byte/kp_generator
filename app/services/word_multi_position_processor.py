# app/services/word_multi_position_processor.py
from io import BytesIO
from typing import Any, Dict, List, Optional

try:
    from docx import Document
except ImportError:
    # Fallback для старых версий python-docx
    try:
        from docx.document import Document
    except ImportError:
        raise ImportError(
            "python-docx не установлен. Установите: pip install python-docx"
        )


class WordMultiPositionProcessor:
    """Обработчик множественных позиций для Word файлов."""
    
    def __init__(self, template_path: str):
        self.template_path = template_path
    
    def process_multiple_positions(
        self,
        positions: List[Dict[str, Any]],
        form_data: Dict[str, Any] = None,
        final_price: float = None,
        general_price: float = None,
        final_price_nds: float = None,
        position_prices: Optional[List[Dict[str, Any]]] = None,
        contact_info: Optional[str] = None,
    ) -> BytesIO:
        """Обрабатывает множественные позиции и возвращает Word файл в памяти."""
        if not positions:
            raise ValueError("Список позиций не может быть пустым")
        
        # Загружаем шаблон
        doc = Document(self.template_path)
        
        # TODO: Реализовать заполнение Word документа данными из positions и form_data
        # Это базовая заглушка, которая просто возвращает загруженный шаблон
        
        # Сохраняем в BytesIO
        buffer = BytesIO()
        doc.save(buffer)
        buffer.seek(0)
        return buffer

