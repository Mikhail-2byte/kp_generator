# Document generation
from datetime import datetime
import zipfile
from io import BytesIO

from app.business.interfaces import DocumentGeneratorPort
from app.presentation.helpers import get_safe_filename, extract_positions_from_form
from app.services.multi_position_processor import MultiPositionProcessor
from app.services.word_multi_position_processor import WordMultiPositionProcessor


class DocumentGenerationService(DocumentGeneratorPort):
    """Сервис генерации документов с возможностью подмены процессоров."""

    def __init__(
        self,
        excel_processor_factory=MultiPositionProcessor,
        word_processor_factory=WordMultiPositionProcessor,
    ):
        self._excel_processor_factory = excel_processor_factory
        self._word_processor_factory = word_processor_factory

    def _resolve_positions(self, form_data, positions):
        resolved = positions or extract_positions_from_form(form_data)
        if not resolved:
            raise ValueError('Список позиций не может быть пустым')
        return resolved

    def generate_excel_document(
        self,
        template_path,
        form_data,
        final_price,
        general_prise,
        position_prices=None,
        manager_fio=None,
    ):
        """Готовит Excel-файл в памяти."""
        positions = self._resolve_positions(form_data, None)
        processor = self._excel_processor_factory(template_path)
        return processor.process_multiple_positions(
            positions,
            form_data,
            final_price,
            general_prise,
            position_prices=position_prices,
            manager_fio=manager_fio,
        )

    def generate_word_document(
        self,
        template_path,
        form_data,
        final_price,
        general_prise,
        final_price_nds,
        positions=None,
        position_prices=None,
        contact_info=None,
    ):
        """Формирует коммерческое предложение в формате Word."""
        resolved_positions = self._resolve_positions(form_data, positions)
        processor = self._word_processor_factory(template_path)
        return processor.process_multiple_positions(
            resolved_positions,
            form_data,
            final_price,
            general_prise,
            final_price_nds=final_price_nds,
            position_prices=position_prices,
            contact_info=contact_info,
        )

    def create_zip_archive(self, excel_file, word_file, company_name):
        """Упаковывает подготовленные документы в ZIP с читаемым именем."""
        timestamp = datetime.now().strftime('%Y%m%d_%H%M')
        file_prefix = f"КП_{get_safe_filename(company_name)}_{timestamp}"

        zip_buffer = BytesIO()
        with zipfile.ZipFile(zip_buffer, 'w', zipfile.ZIP_DEFLATED) as zip_file:
            zip_file.writestr(f"{file_prefix}.xlsx", excel_file.getvalue())
            zip_file.writestr(f"{file_prefix}.docx", word_file.getvalue())

        zip_buffer.seek(0)
        return zip_buffer, file_prefix


_default_generator = DocumentGenerationService()


def generate_excel_document(*args, **kwargs):
    return _default_generator.generate_excel_document(*args, **kwargs)


def generate_word_document(*args, **kwargs):
    return _default_generator.generate_word_document(*args, **kwargs)


def create_zip_archive(*args, **kwargs):
    return _default_generator.create_zip_archive(*args, **kwargs)
