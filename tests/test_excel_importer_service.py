from io import BytesIO

import pytest
from openpyxl import Workbook

from app.services.excel_importer import ExcelImportError, ExcelImporterService


def _build_workbook(header, rows):
    wb = Workbook()
    ws = wb.active
    ws.append(header)
    for row in rows:
        ws.append(row)
    buffer = BytesIO()
    wb.save(buffer)
    buffer.seek(0)
    return buffer


def test_excel_importer_parses_positions():
    header = ['Номенклатура', 'Цена закупа', 'Количество', 'Вес за шт. (кг)', 'Пошлина (%)']
    rows = [
        ['Труба', '1200', '5', '1.5', '7'],
        ['Фланец', '800', '2', '0.8', ''],
    ]
    stream = _build_workbook(header, rows)
    importer = ExcelImporterService()

    positions = importer.parse_positions(stream)

    assert len(positions) == 2
    assert positions[0]['product'] == 'Труба'
    assert positions[1]['duty_percent'] == '0'


def test_excel_importer_missing_columns_raises():
    header = ['Номенклатура', 'Количество']
    rows = [['Труба', '5']]
    stream = _build_workbook(header, rows)
    importer = ExcelImporterService()

    with pytest.raises(ExcelImportError):
        importer.parse_positions(stream)

