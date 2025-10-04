from io import BytesIO

import pytest
from openpyxl import Workbook

from app.services.excel_importer import ExcelImportError, parse_positions_from_excel


def build_workbook(rows) -> BytesIO:
    workbook = Workbook()
    sheet = workbook.active
    sheet.append([
        '№',
        'Номенклатура',
        'Номер чертежа',
        'Материал',
        'Цена закупа',
        'Количество, шт.',
        'Вес за шт. (кг)',
        'Пошлина (%)'
    ])
    for excel_row in rows:
        sheet.append(excel_row)

    buffer = BytesIO()
    workbook.save(buffer)
    buffer.seek(0)
    return buffer


def test_parse_positions_from_excel_success():
    stream = build_workbook([
        [1, 'Вал', 'ЧЧ-01', 'Сталь 40Х', '123.50', 4, 5.5, 12],
        [2, 'Шестерня', '', '', 200, '3', '1,8', None],
    ])

    positions = parse_positions_from_excel(stream)

    assert len(positions) == 2
    assert positions[0]['product'] == 'Вал'
    assert positions[0]['quantity'] == '4'
    assert positions[0]['cost_price'] == '123.5'
    assert positions[1]['quantity'] == '3'
    assert positions[1]['weight'] == '1.8'
    assert positions[1]['duty_percent'] == '0'


def test_parse_positions_from_excel_missing_required_column():
    workbook = Workbook()
    sheet = workbook.active
    sheet.append(['Номенклатура', 'Количество, шт.'])
    sheet.append(['Деталь', 1])

    buffer = BytesIO()
    workbook.save(buffer)
    buffer.seek(0)

    with pytest.raises(ExcelImportError):
        parse_positions_from_excel(buffer)


def test_parse_positions_from_excel_requires_integer_quantity():
    stream = build_workbook([
        [1, 'Деталь', '', '', 100, '1.5', 2, 5],
    ])

    with pytest.raises(ExcelImportError):
        parse_positions_from_excel(stream)
