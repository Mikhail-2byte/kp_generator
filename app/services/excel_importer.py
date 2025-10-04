from __future__ import annotations

from dataclasses import dataclass
from decimal import Decimal, InvalidOperation
from typing import BinaryIO, Dict, Iterable, List

from openpyxl import load_workbook
from openpyxl.utils.exceptions import InvalidFileException


HeaderMap = Dict[str, str]


@dataclass(frozen=True)
class ImportSchema:
    header_map: HeaderMap
    required_fields: Iterable[str]
    numeric_fields: Iterable[str]

    def normalise_header(self, value: object) -> str:
        if value is None:
            return ''
        return str(value).strip().lower()

    def resolve_field(self, header_value: object) -> str | None:
        normalised = self.normalise_header(header_value)
        return self.header_map.get(normalised)


SCHEMA = ImportSchema(
    header_map={
        '№': 'ordinal',
        'no': 'ordinal',
        '#': 'ordinal',
        'номер позиции': 'ordinal',
        'номенклатура': 'product',
        'наименование товара': 'product',
        'наименование продукции': 'product',
        'название': 'product',
        'номер чертежа': 'drawing_number',
        'чертеж': 'drawing_number',
        'материал': 'material',
        'материал изделия': 'material',
        'цена закупа': 'cost_price',
        'цена закупа, руб': 'cost_price',
        'стоимость закупа': 'cost_price',
        'количество, шт.': 'quantity',
        'количество шт.': 'quantity',
        'количество': 'quantity',
        'шт.': 'quantity',
        'вес за шт. (кг)': 'weight',
        'вес за шт. кг': 'weight',
        'вес, кг': 'weight',
        'вес': 'weight',
        'пошлина (%)': 'duty_percent',
        'пошлина %': 'duty_percent',
        'пошлина': 'duty_percent',
    },
    required_fields=('product', 'cost_price', 'quantity', 'weight'),
    numeric_fields=('cost_price', 'quantity', 'weight', 'duty_percent'),
)


FIELD_LABELS = {
    'product': 'Номенклатура',
    'drawing_number': 'Номер чертежа',
    'material': 'Материал',
    'cost_price': 'Цена закупа',
    'quantity': 'Количество, шт.',
    'weight': 'Вес за шт. (кг)',
    'duty_percent': 'Пошлина (%)',
    'ordinal': '№',
}


class ExcelImportError(ValueError):
    """Ошибка валидации Excel при импорте позиций."""


def _normalise_numeric(field: str, value: object, row_idx: int) -> str:
    if value is None:
        return ''

    if isinstance(value, (int, float, Decimal)):
        decimal_value = Decimal(str(value))
    elif isinstance(value, str):
        normalised = value.replace(' ', '').replace(',', '.').strip()
        if not normalised:
            return ''
        try:
            decimal_value = Decimal(normalised)
        except InvalidOperation as exc:
            raise ExcelImportError(
                f"Некорректное числовое значение в столбце '{FIELD_LABELS[field]}' (строка {row_idx})."
            ) from exc
    else:
        raise ExcelImportError(
            f"Некорректный формат данных в столбце '{FIELD_LABELS[field]}' (строка {row_idx})."
        )

    if field == 'quantity':
        if decimal_value != decimal_value.to_integral_value():
            raise ExcelImportError(
                f"Количество в строке {row_idx} должно быть целым числом."
            )
        return str(int(decimal_value))

    normalised_dec = decimal_value.normalize()
    if normalised_dec == normalised_dec.to_integral_value():
        return str(int(normalised_dec))
    return format(normalised_dec, 'f')


def _prepare_value(field: str, value: object, row_idx: int) -> str:
    if field in SCHEMA.numeric_fields:
        return _normalise_numeric(field, value, row_idx)

    if value is None:
        return ''
    return str(value).strip()


def parse_positions_from_excel(stream: BinaryIO) -> List[Dict[str, str]]:
    """Читает Excel шаблон и возвращает список позиций для заполнения формы."""

    if hasattr(stream, 'seek'):
        stream.seek(0)

    try:
        workbook = load_workbook(stream, data_only=True)
    except InvalidFileException as exc:
        raise ExcelImportError('Не удалось прочитать файл Excel. Убедитесь, что используется формат .xlsx.') from exc
    except Exception as exc:  # pragma: no cover - делегирование неожиданных ошибок
        raise ExcelImportError('Файл повреждён или имеет неподдерживаемый формат.') from exc

    sheet = workbook.active

    header_row = None
    for row in sheet.iter_rows(min_row=1, max_row=sheet.max_row):
        if any(str(cell.value).strip() for cell in row if cell.value is not None):
            header_row = row
            break

    if not header_row:
        raise ExcelImportError('Файл не содержит заголовков и не может быть импортирован.')

    column_map: Dict[str, int] = {}
    for idx, cell in enumerate(header_row):
        field_name = SCHEMA.resolve_field(cell.value)
        if field_name and field_name not in column_map:
            column_map[field_name] = idx

    missing = [FIELD_LABELS[field] for field in SCHEMA.required_fields if field not in column_map]
    if missing:
        missing_list = ', '.join(missing)
        raise ExcelImportError(f'В файле отсутствуют обязательные столбцы: {missing_list}.')

    positions: List[Dict[str, str]] = []
    start_row = header_row[0].row + 1

    for row in sheet.iter_rows(min_row=start_row, max_row=sheet.max_row):
        row_idx = row[0].row
        row_payload: Dict[str, str] = {}

        for field, column_index in column_map.items():
            value = row[column_index].value if column_index < len(row) else None
            row_payload[field] = _prepare_value(field, value, row_idx)

        if not any(row_payload.get(field) for field in SCHEMA.required_fields):
            continue

        for field in SCHEMA.required_fields:
            if not row_payload.get(field):
                raise ExcelImportError(
                    f"В строке {row_idx} отсутствует значение в столбце '{FIELD_LABELS[field]}'."
                )

        if 'duty_percent' in row_payload:
            row_payload['duty_percent'] = row_payload.get('duty_percent') or '0'
        else:
            row_payload['duty_percent'] = '0'
        positions.append(row_payload)

    if not positions:
        raise ExcelImportError('Файл не содержит позиций для импорта.')

    return positions


__all__ = ['ExcelImportError', 'parse_positions_from_excel']
