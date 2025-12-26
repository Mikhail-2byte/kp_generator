"""
Обработчик множественных позиций для Word документов.
Аналогичен MultiPositionProcessor для Excel, но работает с python-docx.
"""
from typing import List, Dict, Any, Optional
from docx import Document
from docx.table import Table
from io import BytesIO
from datetime import datetime
import math


class WordMultiPositionProcessor:
    """Обработчик множественных позиций для Word файлов."""
    
    @staticmethod
    def format_price(value: float) -> str:
        """
        Форматирует цену с пробелами в качестве разделителей тысяч.
        
        Args:
            value: Числовое значение цены
            
        Returns:
            Строка с отформатированной ценой (например: "1 234 000", "13 000")
        """
        # Округляем до целого числа
        int_value = int(round(value))
        # Форматируем с запятыми и заменяем на пробелы
        return f"{int_value:,}".replace(",", " ")
    
    @staticmethod
    def format_proposal_validity(value: str) -> str:
        """
        Форматирует срок действия предложения в зависимости от ввода пользователя.
        
        Если введено число (например, "30") - возвращает:
        "Срок действия данного предложения составляет 30 календарных дней"
        
        Если введена дата (например, "12.01.2025") - возвращает:
        "Предложение действительно до 12.01.2025г."
        
        Args:
            value: Введенное пользователем значение
            
        Returns:
            Отформатированная строка или пустая строка, если значение пустое
        """
        if not value or not value.strip():
            return ''
        
        value = value.strip()
        
        # Проверяем, является ли значение числом (только цифры)
        import re
        if re.match(r'^\d+$', value):
            # Это число дней
            days = int(value)
            return f"Срок действия данного предложения составляет {days} календарных дней"
        
        # Проверяем, является ли значение датой (формат DD.MM.YYYY или DD.MM.YYYYг.)
        date_pattern = r'^(\d{1,2})\.(\d{1,2})\.(\d{4})(?:г\.?)?$'
        date_match = re.match(date_pattern, value)
        if date_match:
            # Это дата, форматируем с "г." в конце
            day, month, year = date_match.groups()
            # Убираем возможное "г." из исходного значения и добавляем свое
            formatted_date = f"{day}.{month}.{year}г."
            return f"Предложение действительно до {formatted_date}"
        
        # Если не число и не дата, возвращаем исходное значение
        return value
    
    def __init__(self, template_path: str):
        """
        Инициализирует процессор.
        
        Args:
            template_path: Путь к Word шаблону
        """
        self.template_path = template_path
        self.doc: Optional[Document] = None
        
    def load_template(self) -> Document:
        """Загружает Word шаблон."""
        if self.doc is None:
            self.doc = Document(self.template_path)
        return self.doc
    
    def find_positions_table(self) -> Optional[Table]:
        """
        Находит таблицу с позициями в документе.
        
        Ищет таблицу, которая содержит плейсхолдеры позиций
        (например, {{ product }}, {{ quantity }}, и т.д.).
        
        Returns:
            Таблица с позициями или None, если не найдена
        """
        doc = self.load_template()
        
        position_placeholders = [
            '{{ product }}',
            '{{ quantity }}',
            '{{ cost_price }}',
            '{{ weight }}',
        ]
        
        for table in doc.tables:
            for row in table.rows:
                for cell in row.cells:
                    cell_text = cell.text
                    # Проверяем, есть ли в таблице плейсхолдеры позиций
                    if any(placeholder in cell_text for placeholder in position_placeholders):
                        return table
        
        return None
    
    def find_template_row(self, table: Table) -> Optional[int]:
        """
        Находит индекс строки-шаблона в таблице.
        
        Строка-шаблон - это строка, которая содержит плейсхолдеры позиций
        и будет использоваться для копирования при добавлении новых позиций.
        
        Args:
            table: Таблица для поиска
            
        Returns:
            Индекс строки-шаблона (0-based) или None
        """
        position_placeholders = [
            '{{ product }}',
            '{{ quantity }}',
            '{{ cost_price }}',
        ]
        
        for row_idx, row in enumerate(table.rows):
            row_text = ' '.join(cell.text for cell in row.cells)
            # Проверяем, содержит ли строка плейсхолдеры позиций
            if any(placeholder in row_text for placeholder in position_placeholders):
                return row_idx
        
        return None
    
    def find_total_rows_start(self, table: Table, template_row_idx: int) -> Optional[int]:
        """
        Находит индекс первой строки с итогами (ИТОГО без НДС, ИТОГО с НДС).
        
        Args:
            table: Таблица для поиска
            template_row_idx: Индекс строки-шаблона
            
        Returns:
            Индекс первой строки с итогами или None, если не найдено
        """
        total_keywords = ['ИТОГО', 'Итого', 'итого', 'TOTAL', 'Total', 'total']
        
        for row_idx in range(template_row_idx + 1, len(table.rows)):
            row = table.rows[row_idx]
            row_text = ' '.join(cell.text for cell in row.cells)
            if any(keyword in row_text for keyword in total_keywords):
                return row_idx
        
        return None
    
    def copy_row_formatting(self, source_row: Any, target_row: Any) -> None:
        """
        Копирует форматирование из исходной строки в целевую.
        
        Args:
            source_row: Исходная строка (шаблон)
            target_row: Целевая строка (новая позиция)
        """
        # Копируем форматирование ячеек
        for source_cell, target_cell in zip(source_row.cells, target_row.cells):
            # Очищаем целевую ячейку, но сохраняем структуру
            # Удаляем все параграфы кроме первого
            while len(target_cell.paragraphs) > 1:
                target_cell._element.remove(target_cell.paragraphs[-1]._element)
            
            # Копируем текст и форматирование из исходной ячейки
            if source_cell.paragraphs:
                source_paragraph = source_cell.paragraphs[0]
                
                # Используем первый параграф целевой ячейки или создаем новый
                if target_cell.paragraphs:
                    target_paragraph = target_cell.paragraphs[0]
                    # Очищаем существующие runs
                    for run in target_paragraph.runs:
                        target_paragraph._element.remove(run._element)
                else:
                    target_paragraph = target_cell.add_paragraph()
                
                # Копируем runs из исходного параграфа
                for source_run in source_paragraph.runs:
                    target_run = target_paragraph.add_run(source_run.text)
                    
                    # Копируем свойства форматирования
                    if source_run.font.name:
                        target_run.font.name = source_run.font.name
                    if source_run.font.size:
                        target_run.font.size = source_run.font.size
                    target_run.font.bold = source_run.font.bold
                    target_run.font.italic = source_run.font.italic
                    target_run.font.underline = source_run.font.underline
                
                # Копируем выравнивание параграфа
                if source_paragraph.alignment is not None:
                    target_paragraph.alignment = source_paragraph.alignment
                
                # Копируем форматирование ячейки (если доступно)
                if hasattr(source_cell, 'vertical_alignment'):
                    target_cell.vertical_alignment = source_cell.vertical_alignment
    
    def add_position_row(self, table: Table, template_row_idx: int) -> Any:
        """
        Добавляет новую строку в таблицу, копируя форматирование из строки-шаблона.
        
        Args:
            table: Таблица для добавления строки
            template_row_idx: Индекс строки-шаблона
            
        Returns:
            Новая добавленная строка
        """
        template_row = table.rows[template_row_idx]
        
        # Добавляем новую строку после шаблона
        new_row = table.add_row()
        
        # Копируем форматирование
        self.copy_row_formatting(template_row, new_row)
        
        return new_row
    
    def fill_position_data(self, row: Any, position: Dict[str, Any], position_price: Optional[Dict[str, Any]] = None) -> None:
        """
        Заполняет данные позиции в строке таблицы.
        
        Args:
            row: Строка таблицы для заполнения
            position: Данные позиции
            position_price: Рассчитанные цены позиции (опционально)
        """
        # Вспомогательная функция для безопасного преобразования в число
        def safe_float(value, default=0.0):
            if value is None:
                return default
            try:
                return float(value)
            except (ValueError, TypeError):
                return default
        
        def safe_int(value, default=0):
            if value is None:
                return default
            try:
                return int(float(value))
            except (ValueError, TypeError):
                return default
        
        # Маппинг полей позиции на плейсхолдеры
        # Обрабатываем строковые и числовые значения
        position_data = {
            '{{ product }}': str(position.get('product', '')).strip(),
            '{{ drawing_number }}': str(position.get('drawing_number', '')).strip(),
            '{{ material }}': str(position.get('material', '')).strip(),
            '{{ quantity }}': str(safe_int(position.get('quantity', 0))),
            '{{ cost_price }}': f"{safe_float(position.get('cost_price', 0)):.0f}",
            '{{ weight }}': f"{safe_float(position.get('weight', 0)):.0f}",
            '{{ duty_percent }}': f"{safe_float(position.get('duty_percent', 0)):.1f}",
        }
        
        # Добавляем рассчитанные цены, если они переданы
        # Важно: цены должны быть рассчитаны и переданы в position_price
        if position_price and isinstance(position_price, dict):
            final_price = safe_float(position_price.get('final_price', 0))
            general_price = safe_float(position_price.get('general_price', 0))
            # Округляем цену за единицу вверх до целого числа (чтобы маржа была не меньше целевой)
            final_price_rounded = math.ceil(final_price) if final_price > 0 else 0
            # Форматируем цены с пробелами в качестве разделителей тысяч
            # Всегда добавляем final_price, даже если 0 (чтобы заменить плейсхолдер)
            position_data['{{ final_price }}'] = self.format_price(final_price_rounded)
            # Всегда добавляем general_price для позиции
            # Пересчитываем общую цену с учетом округленной цены за единицу
            quantity = safe_int(position.get('quantity', 0))
            general_price_recalculated = final_price_rounded * quantity if quantity > 0 else general_price
            position_data['{{ position_total }}'] = self.format_price(general_price_recalculated)
            # Также поддерживаем старый плейсхолдер для обратной совместимости
            position_data['{{ general_prise }}'] = self.format_price(general_price_recalculated)
        else:
            # Если position_price не передан, устанавливаем значения по умолчанию
            # чтобы плейсхолдеры все равно заменялись
            position_data['{{ final_price }}'] = "0"
            position_data['{{ position_total }}'] = "0"
            position_data['{{ general_prise }}'] = "0"
        
        # Заменяем плейсхолдеры в ячейках строки
        for cell in row.cells:
            # Получаем полный текст ячейки со всеми параграфами
            cell_text = cell.text
            original_text = cell_text
            
            # Заменяем все плейсхолдеры (даже если есть символы вокруг, например ¥{{ final_price }})
            # Важно: заменяем все плейсхолдеры, включая {{ final_price }}
            for placeholder, value in position_data.items():
                if placeholder in cell_text:
                    # Заменяем плейсхолдер на значение
                    cell_text = cell_text.replace(placeholder, value)
            
            # Проверяем, были ли изменения
            if cell_text != original_text:
                
                # Обновляем текст ячейки
                # Очищаем ячейку и устанавливаем новый текст
                # Сохраняем форматирование первого параграфа, если возможно
                if cell.paragraphs:
                    paragraph = cell.paragraphs[0]
                    # Сохраняем форматирование первого run
                    if paragraph.runs:
                        first_run = paragraph.runs[0]
                        # Сохраняем форматирование
                        font_name = first_run.font.name if first_run.font.name else None
                        font_size = first_run.font.size
                        is_bold = first_run.font.bold
                        is_italic = first_run.font.italic
                        
                        # Очищаем все runs
                        for run in paragraph.runs[:]:
                            paragraph._element.remove(run._element)
                        
                        # Создаем новый run с сохраненным форматированием
                        new_run = paragraph.add_run(cell_text)
                        if font_name:
                            new_run.font.name = font_name
                        if font_size:
                            new_run.font.size = font_size
                        new_run.font.bold = is_bold
                        new_run.font.italic = is_italic
                    else:
                        # Если нет runs, просто устанавливаем текст
                        paragraph.text = cell_text
                else:
                    # Если нет параграфов, создаем новый
                    cell.add_paragraph(cell_text)
    
    def replace_common_placeholders(self, form_data: Dict[str, Any], final_price: float = None, 
                                   general_price: float = None, final_price_nds: float = None, 
                                   contact_info: str = None) -> None:
        """
        Заменяет общие плейсхолдеры в документе (компания, дата, итоговые суммы и т.д.).
        
        Args:
            form_data: Данные формы
            final_price: Цена за единицу (опционально)
            general_price: Общая цена (опционально)
            final_price_nds: Цена с НДС (опционально)
            contact_info: Контактная информация менеджера (опционально)
        """
        doc = self.load_template()
        
        current_date = datetime.now().strftime('%d.%m.%Yг.')
        company = form_data.get('company', '').strip()
        logistics = float(form_data.get('logistics', 0))
        delivery_time = int(form_data.get('delivery_time', 0))
        tender_number = form_data.get('tender_number', '').strip()
        delivery_address = form_data.get('delivery_address', '').strip()
        payment_terms = form_data.get('payment_terms', '').strip()
        proposal_validity_raw = form_data.get('proposal_validity', '').strip()
        warranty_period_raw = form_data.get('warranty_period', '').strip()
        # Если гарантийный срок не указан, используем значение по умолчанию
        if not warranty_period_raw:
            warranty_period_raw = '12 месяцев со дня ввода в эксплуатацию, но не более 18 месяцев со дня поставки.'
        # Форматируем срок действия предложения (число -> текст с днями, дата -> текст с датой)
        proposal_validity = self.format_proposal_validity(proposal_validity_raw)
        # Форматируем гарантийный срок с префиксом
        warranty_period = f"Гарантийный срок эксплуатации изделия - {warranty_period_raw}"
        
        word_data = {
            '{{ company }}': company,
            '{{ logistics }}': f"{logistics:.0f}",
            '{{ delivery_time }}': str(delivery_time),
            '{{ tender_number }}': tender_number,
            '{{ delivery_address }}': delivery_address,
            '{{ payment_terms }}': payment_terms,
            '{{ proposal_validity }}': proposal_validity,
            '{{ warranty_period }}': warranty_period,
            '{{ date }}': current_date,
        }
        
        # Всегда добавляем контактную информацию (заменяем на пустоту, если пользователь не авторизован или не указал контакты)
        if contact_info:
            word_data['{{ contact_info }}'] = contact_info.strip()
        else:
            word_data['{{ contact_info }}'] = ''
        
        # Добавляем итоговые цены, если они переданы
        # ВАЖНО: НЕ добавляем {{ final_price }} и {{ general_prise }} в word_data,
        # так как они будут заполнены индивидуально для каждой позиции в таблице
        # Эти плейсхолдеры используются только в футере и других местах, не в строках позиций
        if final_price_nds is not None:
            word_data['{{ final_price_NDS }}'] = self.format_price(final_price_nds)
        # Добавляем общую цену только для футера (не для строк позиций)
        if general_price is not None:
            word_data['{{ general_prise }}'] = self.format_price(general_price)
        
        # Заменяем плейсхолдеры в параграфах
        for paragraph in doc.paragraphs:
            for key, value in word_data.items():
                if key in paragraph.text:
                    paragraph.text = paragraph.text.replace(key, value)
        
        # Заменяем плейсхолдеры в таблицах (кроме таблицы с позициями)
        # ВАЖНО: не заменяем {{ final_price }} и {{ general_prise }} в таблице с позициями,
        # так как они будут заполнены индивидуально для каждой позиции
        positions_table = self.find_positions_table()
        template_row_idx = None
        if positions_table:
            template_row_idx = self.find_template_row(positions_table)
        
        for table in doc.tables:
            if table != positions_table:  # Пропускаем таблицу с позициями
                for row in table.rows:
                    for cell in row.cells:
                        for key, value in word_data.items():
                            if key in cell.text:
                                cell.text = cell.text.replace(key, value)
            else:
                # Для таблицы с позициями заменяем плейсхолдеры только в футере
                # (строках после строки-шаблона, которые не являются строками позиций)
                if template_row_idx is not None:
                    # Заменяем плейсхолдеры только в строках после всех позиций (в футере)
                    # Строки позиций будут заполнены позже индивидуально
                    for row_idx, row in enumerate(table.rows):
                        # Пропускаем строки позиций (до и включая строку-шаблон + потенциальные позиции)
                        # Заменяем только в футере (после строки-шаблона + несколько строк для позиций)
                        # Для безопасности заменяем только в строках, которые явно не являются строками позиций
                        # (например, строки с "ИТОГО")
                        row_text = ' '.join(cell.text for cell in row.cells)
                        if 'ИТОГО' in row_text or 'итого' in row_text.lower():
                            # Это строка футера, заменяем плейсхолдеры
                            for cell in row.cells:
                                for key, value in word_data.items():
                                    if key in cell.text:
                                        cell.text = cell.text.replace(key, value)
    
    def process_multiple_positions(
        self,
        positions: List[Dict[str, Any]],
        form_data: Dict[str, Any] = None,
        final_price: float = None,
        general_price: float = None,
        final_price_nds: float = None,
        position_prices: List[Dict[str, Any]] = None,
        contact_info: str = None
    ) -> BytesIO:
        """
        Обрабатывает множественные позиции и возвращает Word файл в памяти.
        
        Args:
            positions: Список позиций
            form_data: Данные формы (опционально)
            final_price: Цена за единицу (опционально)
            general_price: Общая цена (опционально)
            final_price_nds: Цена с НДС (опционально)
            position_prices: Список рассчитанных цен для каждой позиции (опционально)
            contact_info: Контактная информация менеджера (опционально)
            
        Returns:
            BytesIO объект с Word документом
        """
        if not positions:
            raise ValueError("Список позиций не может быть пустым")
        
        # Загружаем шаблон
        doc = self.load_template()
        
        # Заменяем общие плейсхолдеры
        if form_data:
            self.replace_common_placeholders(form_data, final_price, general_price, final_price_nds, contact_info)
        
        # Находим таблицу с позициями
        positions_table = self.find_positions_table()
        if not positions_table:
            raise ValueError("Не найдена таблица с позициями в шаблоне")
        
        # Находим строку-шаблон
        template_row_idx = self.find_template_row(positions_table)
        if template_row_idx is None:
            raise ValueError("Не найдена строка-шаблон с плейсхолдерами позиций")
        
        # Находим строки с итогами (чтобы вставлять новые строки ПЕРЕД ними)
        total_row_start = self.find_total_rows_start(positions_table, template_row_idx)
        
        # Количество позиций для добавления (минус 1, так как первая строка уже есть)
        positions_to_add = len(positions) - 1
        
        # Определяем, куда вставлять новые строки
        # Если есть итоговые строки, вставляем перед ними, иначе после шаблона
        if total_row_start is not None:
            insert_before_idx = total_row_start
        else:
            insert_before_idx = len(positions_table.rows)
        
        # Добавляем дополнительные строки, если нужно
        # Вставляем их перед итоговыми строками используя прямое манипулирование XML
        template_row = positions_table.rows[template_row_idx]
        
        if positions_to_add > 0:
            # Получаем XML элемент строки-шаблона
            template_row_xml = template_row._tr
            
            # Находим индекс, куда вставлять (перед итоговыми строками)
            if total_row_start is not None:
                # Вставляем перед итоговыми строками
                target_row_xml = positions_table.rows[total_row_start]._tr
            else:
                # Вставляем в конец (перед последней строкой или в конец)
                target_row_xml = None
            
            # Создаем новые строки, копируя XML из шаблона
            # Вставляем их в обратном порядке, чтобы они оказались в правильном порядке
            # (addprevious() вставляет перед целевым элементом, поэтому последняя вставленная
            # строка будет первой перед целевым элементом)
            new_rows_xml = []
            for _ in range(positions_to_add):
                # Клонируем XML элемент строки-шаблона
                new_row_xml = template_row_xml.__copy__()
                new_rows_xml.append(new_row_xml)
            
            # Вставляем строки в обратном порядке, чтобы они оказались в правильном порядке
            if target_row_xml is not None:
                for new_row_xml in reversed(new_rows_xml):
                    # Вставляем перед целевой строкой используя addprevious()
                    target_row_xml.addprevious(new_row_xml)
            else:
                # Добавляем в конец таблицы
                tbl = positions_table._tbl
                for new_row_xml in new_rows_xml:
                    tbl.append(new_row_xml)
        
        # Пересоздаем список строк после вставки (python-docx обновит индексы)
        # Теперь заполняем данные для всех позиций
        for i, position in enumerate(positions):
            if i == 0:
                # Первая позиция - используем строку-шаблон
                row = positions_table.rows[template_row_idx]
            else:
                # Остальные позиции - используем новые строки
                # После вставки новых строк перед итогами, их индексы = template_row_idx + i
                # Но если есть итоговые строки, нужно учесть их смещение
                if total_row_start is not None:
                    # Новые строки вставлены перед итогами, поэтому их индексы = template_row_idx + i
                    row_idx = template_row_idx + i
                else:
                    # Новые строки добавлены в конец
                    row_idx = template_row_idx + i
                
                if row_idx < len(positions_table.rows):
                    row = positions_table.rows[row_idx]
                else:
                    # Если индекс выходит за границы, используем последнюю строку
                    row = positions_table.rows[-1]
            
            # Получаем рассчитанные цены для позиции, если они переданы
            position_price = None
            if position_prices and i < len(position_prices):
                position_price = position_prices[i]
            
            # Заполняем данные позиции
            self.fill_position_data(row, position, position_price)
        
        # Заполняем итоговые значения в футере таблицы с позициями (если есть)
        # Ищем строки после всех позиций, которые могут содержать итоговые плейсхолдеры
        # После вставки новых строк, итоговые строки сдвинулись
        if total_row_start is not None:
            # Итоговые строки были найдены, их индекс сдвинулся на количество добавленных строк
            updated_total_row_start = total_row_start + positions_to_add
        else:
            # Итоговых строк не было, ищем после всех позиций
            updated_total_row_start = template_row_idx + len(positions)
        
        for row_idx in range(updated_total_row_start, len(positions_table.rows)):
            row = positions_table.rows[row_idx]
            for cell in row.cells:
                cell_text = cell.text
                original_text = cell_text
                
                # Заменяем общие итоговые плейсхолдеры в футере
                # Учитываем, что могут быть символы вокруг (например ¥{{ general_prise }})
                # В футере {{ general_prise }} - это общая цена ВСЕХ позиций (без НДС)
                # В строках позиций {{ general_prise }} или {{ position_total }} - это цена ОДНОЙ позиции
                if '{{ general_prise }}' in cell_text and general_price is not None:
                    # В футере заменяем на общую цену всех позиций с форматированием
                    cell_text = cell_text.replace('{{ general_prise }}', self.format_price(general_price))
                if '{{ final_price_NDS }}' in cell_text and final_price_nds is not None:
                    cell_text = cell_text.replace('{{ final_price_NDS }}', self.format_price(final_price_nds))
                
                # Обновляем текст ячейки, если были изменения
                if cell_text != original_text:
                    if cell.paragraphs:
                        paragraph = cell.paragraphs[0]
                        # Сохраняем форматирование первого run
                        if paragraph.runs:
                            first_run = paragraph.runs[0]
                            # Сохраняем форматирование
                            font_name = first_run.font.name if first_run.font.name else None
                            font_size = first_run.font.size
                            is_bold = first_run.font.bold
                            is_italic = first_run.font.italic
                            
                            # Очищаем все runs
                            for run in paragraph.runs[:]:
                                paragraph._element.remove(run._element)
                            
                            # Создаем новый run с сохраненным форматированием
                            new_run = paragraph.add_run(cell_text)
                            if font_name:
                                new_run.font.name = font_name
                            if font_size:
                                new_run.font.size = font_size
                            new_run.font.bold = is_bold
                            new_run.font.italic = is_italic
                        else:
                            paragraph.text = cell_text
                    else:
                        cell.add_paragraph(cell_text)
        
        # Сохраняем в BytesIO
        word_file = BytesIO()
        doc.save(word_file)
        word_file.seek(0)
        
        return word_file

