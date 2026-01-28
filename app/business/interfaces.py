from __future__ import annotations

from io import BytesIO
from typing import BinaryIO, Dict, List, Protocol, Tuple


class PriceCalculatorPort(Protocol):
    """Интерфейс калькуляторов стоимости."""

    def calculate_positions(
        self,
        positions: List[Dict[str, object]],
        logistics_rub: float,
        delivery_time: int,
        margin_percent: float,
    ) -> Dict[str, object]:
        """Возвращает расчёт итоговых цен для всех позиций."""


class DocumentGeneratorPort(Protocol):
    """Интерфейс генераторов документов и архивов."""

    def generate_excel_document(
        self,
        template_path: str,
        form_data: Dict[str, object],
        final_price: float,
        general_price: float,
        positions: List[Dict[str, object]] | None = None,
        position_prices: List[Dict[str, object]] | None = None,
        manager_fio: str | None = None,
    ) -> BytesIO:
        """Готовит Excel-файл на основе шаблона."""

    def generate_word_document(
        self,
        template_path: str,
        form_data: Dict[str, object],
        final_price: float,
        general_price: float,
        final_price_nds: float,
        positions: List[Dict[str, object]] | None = None,
        position_prices: List[Dict[str, object]] | None = None,
        contact_info: str | None = None,
    ) -> BytesIO:
        """Генерирует Word-документ коммерческого предложения."""

    def create_zip_archive(
        self,
        excel_file: BytesIO,
        word_file: BytesIO,
        company_name: str,
    ) -> Tuple[BytesIO, str]:
        """Упаковывает документы в архив и возвращает буфер и префикс имени."""


class ExcelImporterPort(Protocol):
    """Интерфейс импорта позиций из Excel."""

    def parse_positions(self, stream: BinaryIO) -> List[Dict[str, str]]:
        """Возвращает список позиций, считанных из Excel файла."""


__all__ = ['PriceCalculatorPort', 'DocumentGeneratorPort', 'ExcelImporterPort']

