import io
import zipfile

import pytest

from app.business.document_generator import DocumentGenerationService


class _DummyProcessor:
    def __init__(self, template_path):
        self.template_path = template_path

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

    assert prefix.startswith('КП_')
    with zipfile.ZipFile(zip_buffer) as archive:
        assert any(name.endswith('.xlsx') for name in archive.namelist())
        assert any(name.endswith('.docx') for name in archive.namelist())


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

