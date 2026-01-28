import io
import sys
import zipfile
from pathlib import Path

# Добавляем корень проекта в путь для возможности прямого запуска
project_root = Path(__file__).resolve().parents[2]
sys.path.insert(0, str(project_root))

import pytest

from app.business.document_generator import DocumentGenerationService


class _DummyProcessor:
    def __init__(self, template_path, config=None):
        self.template_path = template_path
        self.config = config

    def process_multiple_positions(self, *args, **kwargs):
        payload = {
            'template_path': self.template_path,
            'args': args,
            'kwargs': kwargs,
        }
        buffer = io.BytesIO()
        buffer.write(str(payload).encode('utf-8'))
        buffer.seek(0)
        return buffer


def _form_data():
    return {
        'product': 'Изделие X',
        'quantity': '2',
        'cost_price': '1500',
        'weight': '3',
        'duty_percent': '5',
    }


def test_document_generation_service_excel(monkeypatch):
    service = DocumentGenerationService(
        excel_processor_factory=_DummyProcessor,
        word_processor_factory=_DummyProcessor,
    )

    excel = service.generate_excel_document(
        'template.xlsx',
        _form_data(),
        final_price=2000,
        general_prise=4000,
        position_prices=[{'final_price': 2000}],
        manager_fio='Иванов Иван',
    )

    assert excel.getvalue()


def test_document_generation_service_word_with_positions():
    service = DocumentGenerationService(
        excel_processor_factory=_DummyProcessor,
        word_processor_factory=_DummyProcessor,
    )

    positions = [{'product': 'Изделие X', 'quantity': '2', 'cost_price': '1000', 'weight': '1'}]
    word = service.generate_word_document(
        'template.docx',
        _form_data(),
        final_price=2100,
        general_prise=4200,
        final_price_nds=5040,
        positions=positions,
        contact_info='test@example.com',
    )
    assert word.getvalue()


def test_document_generation_service_zip():
    service = DocumentGenerationService()
    excel = io.BytesIO(b'excel-bytes')
    word = io.BytesIO(b'word-bytes')
    zip_buffer, prefix = service.create_zip_archive(excel, word, 'ООО Ромашка')

    # Новый формат: "company" или "tender_number - company - margin%"
    # Если нет tender_number и margin, то просто "company"
    assert prefix == 'ООО_Ромашка' or prefix.startswith('КП_')
    with zipfile.ZipFile(zip_buffer) as archive:
        names = archive.namelist()
        assert any(name.endswith('.xlsx') for name in names)
        assert any(name.endswith('.docx') for name in names)


def test_document_generation_service_zip_with_tender_and_margin():
    service = DocumentGenerationService()
    excel = io.BytesIO(b'excel-bytes')
    word = io.BytesIO(b'word-bytes')

    tender_number = '12345'
    margin = '25'

    zip_buffer, prefix = service.create_zip_archive(
        excel,
        word,
        'ООО Ромашка',
        tender_number=tender_number,
        margin_percent=margin,
    )

    # Имя архива должно содержать номер тендера, компанию и маржу с %
    assert '12345' in prefix
    assert 'ООО_Ромашка' in prefix
    assert '25%' in prefix

    with zipfile.ZipFile(zip_buffer) as archive:
        names = archive.namelist()
        # Внутри архива должны быть файлы с фиксированными именами
        assert 'КП.docx' in names
        assert f'Бюджет {tender_number}.xlsx' in names


def test_document_generation_service_empty_positions_error():
    service = DocumentGenerationService(
        excel_processor_factory=_DummyProcessor,
        word_processor_factory=_DummyProcessor,
    )

    with pytest.raises(ValueError):
        service.generate_excel_document(
            'template.xlsx',
            {},
            final_price=1000,
            general_prise=1000,
        )


def test_fill_position_data_writes_duty_percent_zero():
    """Пошлина 0% должна записываться в ячейку Excel (столбец X)."""
    import openpyxl
    from app.services.multi_position_processor import MultiPositionProcessor

    wb = openpyxl.Workbook()
    sheet = wb.active
    processor = MultiPositionProcessor('dummy.xlsx', config={})
    row_number = 10
    position_data = {
        'product': 'Товар без пошлины',
        'quantity': '5',
        'cost_price': '500',
        'weight': '2',
        'duty_percent': '0',
    }
    processor.fill_position_data(sheet, position_data, row_number)
    cell_x = sheet[f'X{row_number}']
    assert cell_x.value is not None
    assert cell_x.value == 0.0


if __name__ == "__main__":
    # Запуск тестов при прямом выполнении файла
    pytest.main([__file__, "-v"])

