# app/document_generator.py
from docx import Document
from io import BytesIO
from datetime import datetime
import zipfile
from app.helpers import get_safe_filename, extract_positions_from_form
from app.services.multi_position_processor import MultiPositionProcessor
from app.services.word_multi_position_processor import WordMultiPositionProcessor


def generate_excel_document(
    template_path,
    form_data,
    final_price,
    general_prise,
    position_prices=None,
    manager_fio=None,
):
    """Готовит Excel-файл в памяти."""
    # Извлекаем позиции из формы
    positions = extract_positions_from_form(form_data)

    if not positions:
        raise ValueError("Список позиций не может быть пустым")

    processor = MultiPositionProcessor(template_path)
    return processor.process_multiple_positions(
        positions,
        form_data,
        final_price,
        general_prise,
        position_prices=position_prices,
        manager_fio=manager_fio,
    )


def generate_word_document(
    template_path,
    form_data,
    final_price,
    general_prise,
    final_price_NDS,
    positions=None,
    position_prices=None,
    contact_info=None,
):
    """
    Формирует коммерческое предложение в формате Word на основе шаблона.
    
    Поддерживает как одну, так и множественные позиции.
    Если positions не передан, извлекает позиции из form_data (обратная совместимость).
    
    Args:
        template_path: Путь к Word шаблону
        form_data: Данные формы
        final_price: Цена за единицу
        general_prise: Общая цена
        final_price_NDS: Цена с НДС
        positions: Список позиций (опционально, для множественных позиций)
        position_prices: Список рассчитанных цен для каждой позиции (опционально)
        contact_info: Контактная информация менеджера (опционально)
    """
    # Извлекаем позиции, если они не переданы (обратная совместимость)
    if positions is None:
        positions = extract_positions_from_form(form_data)
    
    if not positions:
        raise ValueError("Список позиций не может быть пустым")
    
    # Используем новый процессор для множественных позиций
    processor = WordMultiPositionProcessor(template_path)
    return processor.process_multiple_positions(
        positions,
        form_data,
        final_price,
        general_prise,
        final_price_nds=final_price_NDS,
        position_prices=position_prices,
        contact_info=contact_info,
    )


def create_zip_archive(excel_file, word_file, company_name):
    """Упаковывает подготовленные документы в ZIP с читаемым именем."""
    timestamp = datetime.now().strftime('%Y%m%d_%H%M')
    file_prefix = f"КП_{get_safe_filename(company_name)}_{timestamp}"

    zip_buffer = BytesIO()
    with zipfile.ZipFile(zip_buffer, 'w', zipfile.ZIP_DEFLATED) as zip_file:
        zip_file.writestr(f"{file_prefix}.xlsx", excel_file.getvalue())
        zip_file.writestr(f"{file_prefix}.docx", word_file.getvalue())

    zip_buffer.seek(0)
    return zip_buffer, file_prefix
