# app/document_generator.py
from docx import Document
from io import BytesIO
from datetime import datetime
import zipfile
from app.helpers import get_safe_filename, extract_positions_from_form
from app.services.multi_position_processor import MultiPositionProcessor


def generate_excel_document(
    template_path,
    form_data,
    final_price,
    general_prise,
    position_prices=None,
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
    )


def generate_word_document(
    template_path,
    form_data,
    final_price,
    general_prise,
    final_price_NDS,
):
    """Формирует коммерческое предложение в формате Word на основе шаблона."""
    doc = Document(template_path)

    current_date = datetime.now().strftime('%d.%m.%Yг.')
    company = form_data['company'].strip()
    product = form_data['product'].strip()
    quantity = int(form_data['quantity'])
    cost_price = float(form_data['cost_price'])
    weight = float(form_data['weight'])
    logistics = float(form_data['logistics'])
    delivery_time = int(form_data['delivery_time'])
    tender_number = form_data.get('tender_number', '').strip()
    drawing_number = form_data.get('drawing_number', '').strip()
    material = form_data.get('material', '').strip()
    delivery_address = form_data.get('delivery_address', '').strip()
    duty_percent = float(form_data.get('duty_percent', 0))

    word_data = {
        '{{ company }}': company,
        '{{ product }}': product,
        '{{ quantity }}': str(quantity),
        '{{ cost_price }}': f"{cost_price:.0f}",
        '{{ weight }}': f"{weight:.0f}",
        '{{ logistics }}': f"{logistics:.0f}",
        '{{ final_price }}': f"{final_price:.0f}",
        '{{ general_prise }}': f"{general_prise:.0f}",  # Общая цена
        '{{ final_price_NDS }}': f"{final_price_NDS:.0f}",
        '{{ tender_number }}': tender_number,
        '{{ drawing_number }}': drawing_number,
        '{{ material }}': material,
        '{{ delivery_address }}': delivery_address,
        '{{ date }}': current_date,
        '{{ duty_percent }}': f"{duty_percent:.1f}",
        '{{ delivery_time }}': str(delivery_time),
    }

    for paragraph in doc.paragraphs:
        for key, value in word_data.items():
            if key in paragraph.text:
                paragraph.text = paragraph.text.replace(key, value)

    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for key, value in word_data.items():
                    if key in cell.text:
                        cell.text = cell.text.replace(key, value)

    word_file = BytesIO()
    doc.save(word_file)
    word_file.seek(0)
    return word_file


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
