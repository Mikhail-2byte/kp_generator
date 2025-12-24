from app.services import datasets


def test_load_tnved_catalog_uses_static_csv_if_available():
    """
    Проверяем, что загрузка каталога ТН ВЭД не падает
    и возвращает список элементов, когда CSV лежит в static/.

    В реальном проекте файл
    static/ТН-ВЭД-ТД-для-менеджеров-с-ключевыми-словами.csv
    присутствует в репозитории, поэтому функция должна найти его
    через _tnved_catalog_path().
    """
    items = datasets.load_tnved_catalog()
    # Функция не должна возвращать None/исключения
    assert isinstance(items, list)

