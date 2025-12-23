import io
import zipfile


def test_generate_single_position_success(client):
    response = client.post(
        '/generate',
        data={
            'company': 'ООО Тест',
            'delivery_time': '30',
            'logistics': '10000',
            'margin_percent': '20',
            'tender_number': 'TN-001',
            'product': 'Изделие 1',
            'quantity': '5',
            'cost_price': '1200',
            'weight': '10',
            'duty_percent': '5',
            'comment': '',
        },
    )

    assert response.status_code == 200
    assert response.headers.get('Content-Type') == 'application/zip'

    # Проверяем имя скачиваемого ZIP-файла
    content_disposition = response.headers.get('Content-Disposition', '')
    # Имя должно содержать номер тендера, компанию и маржу с процента
    assert 'TN-001' in content_disposition
    assert 'ООО_Тест' in content_disposition or 'ООО Тест' in content_disposition
    assert '20%' in content_disposition

    zip_bytes = io.BytesIO(response.data)
    with zipfile.ZipFile(zip_bytes) as archive:
        names = archive.namelist()
        # Точные имена файлов внутри архива
        assert 'КП.docx' in names
        assert 'Бюджет TN-001.xlsx' in names

