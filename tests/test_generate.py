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

    zip_bytes = io.BytesIO(response.data)
    with zipfile.ZipFile(zip_bytes) as archive:
        names = archive.namelist()
        assert any(name.endswith('.xlsx') for name in names)
        assert any(name.endswith('.docx') for name in names)

