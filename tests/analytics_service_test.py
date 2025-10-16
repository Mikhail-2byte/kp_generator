from pathlib import Path

from werkzeug.datastructures import FileStorage

from app.services.analytics_service import analyze_excel


def test_analyze_excel_product():
    sample_path = Path(r"D:\kp_generator\Analizator\test2.xls")
    with sample_path.open('rb') as file_handle:
        storage = FileStorage(stream=file_handle, filename=sample_path.name)
        result = analyze_excel(storage)

    assert result['analysis_type'] in {'product', 'tender'}
    assert result['metrics']
    assert result['preview'].rows
    assert isinstance(result['download']['csv'], str)
