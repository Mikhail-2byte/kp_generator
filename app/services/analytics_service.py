import base64
from dataclasses import dataclass
from io import BytesIO
from typing import Any, Dict, List, Optional

import matplotlib

matplotlib.use('Agg')
import matplotlib.pyplot as plt
import numpy as np
import pandas as pd
import seaborn as sns


sns.set_theme(style='whitegrid')


class AnalyticsProcessingError(Exception):
    """Ошибка при обработке файла аналитики."""


@dataclass
class PreviewTable:
    columns: List[str]
    rows: List[Dict[str, Any]]


def _fig_to_base64(fig: plt.Figure) -> str:
    buffer = BytesIO()
    fig.savefig(buffer, format='png', bbox_inches='tight')
    plt.close(fig)
    buffer.seek(0)
    return base64.b64encode(buffer.read()).decode('utf-8')


def _normalize_numeric(series: pd.Series) -> pd.Series:
    if series.dtype == object:
        cleaned = (
            series.astype(str)
            .str.replace('\xa0', '', regex=False)
            .str.replace(' ', '', regex=False)
            .str.replace(',', '.', regex=False)
        )
        return pd.to_numeric(cleaned, errors='coerce')
    return pd.to_numeric(series, errors='coerce')


def detect_data_type(df_raw: pd.DataFrame) -> str:
    sample_text = df_raw.astype(str).iloc[:10].to_string()
    tender_keywords = ['Мониторинг', 'участвовали', 'Подались', 'Проиграли', 'Отменен', 'тендер', 'Тендер']
    product_keywords = ['ч.', 'заказчик', 'вес', 'цена', 'количество', 'материал']

    tender_score = sum(1 for keyword in tender_keywords if keyword.lower() in sample_text.lower())
    product_score = sum(1 for keyword in product_keywords if keyword.lower() in sample_text.lower())
    return 'tender' if tender_score >= product_score else 'product'


def process_tender_data(df_raw: pd.DataFrame) -> Optional[pd.DataFrame]:
    sections: List[tuple[str, int]] = []
    for idx, row in df_raw.iterrows():
        cell = str(row.iloc[0]) if len(row) else ''
        if pd.notna(row.iloc[0]) and any(keyword in cell for keyword in ['Мониторинг', 'участвовали', 'Отменен', 'Подались', 'Проиграли', 'Итого']):
            sections.append((cell, idx))

    all_data: List[pd.DataFrame] = []
    for i, (section_name, start_idx) in enumerate(sections):
        end_idx = sections[i + 1][1] if i + 1 < len(sections) else len(df_raw)
        section_data = df_raw.iloc[start_idx:end_idx].copy()

        header_row = None
        for j in range(len(section_data)):
            value = str(section_data.iloc[j, 0])
            if pd.notna(section_data.iloc[j, 0]) and value.isdigit():
                header_row = j
                break

        if header_row is None:
            continue

        headers: List[str] = []
        for col in range(section_data.shape[1]):
            header_val = section_data.iloc[header_row - 1, col] if header_row > 0 else f'Column_{col}'
            if pd.isna(header_val) or str(header_val).strip() == '':
                headers.append(f'Column_{col}')
            else:
                headers.append(str(header_val).strip())

        data_part = section_data.iloc[header_row:].copy()
        data_part.columns = headers
        data_part['Раздел'] = section_name
        first_col = headers[0]
        data_part = data_part.dropna(subset=[first_col])
        data_part = data_part[data_part[first_col].astype(str).str.isdigit()]

        if not data_part.empty:
            all_data.append(data_part)

    if not all_data:
        return None

    df = pd.concat(all_data, ignore_index=True)
    column_mapping = {
        0: '№',
        1: 'Дата начала',
        2: 'Дата окончания',
        3: 'Номер заявки',
        4: 'Заказчик',
        5: 'Тендер',
        6: 'Признак',
        7: 'Тип раздела',
        8: 'Площадка',
        9: 'Комментарий',
        10: 'Сумма',
        11: 'Количество'
    }
    new_columns: List[str] = []
    for idx, col in enumerate(df.columns):
        new_columns.append(column_mapping.get(idx, f'Column_{idx}'))
    df.columns = new_columns
    df = df.replace('', np.nan)

    for col in ['Дата начала', 'Дата окончания']:
        if col in df.columns:
            df[col] = pd.to_datetime(df[col], errors='coerce', dayfirst=True)

    if 'Сумма' in df.columns:
        df['Сумма'] = _normalize_numeric(df['Сумма'])
    if 'Количество' in df.columns:
        df['Количество'] = _normalize_numeric(df['Количество'])
    return df


def process_product_data(df_raw: pd.DataFrame) -> pd.DataFrame:
    df = df_raw.copy()
    headers = [
        'Наименование', 'Чертеж', 'Заказчик', 'Количество', 'Вес_кг', 'Материал',
        'Цена', 'Цена_за_кг', 'Цена_за_кг_с_НДС', 'Менеджер', 'Дата_статус',
        'Тендер', 'Площадка', 'Суммарный_вес', 'Цена_за_кг_итого'
    ]

    if df.shape[1] > len(headers):
        df = df.iloc[:, :len(headers)]
    elif df.shape[1] < len(headers):
        for i in range(df.shape[1], len(headers)):
            df[i] = np.nan

    df.columns = headers

    df = df[df['Чертеж'].notna() & (df['Чертеж'].astype(str).str.strip() != '')]

    numeric_cols = ['Количество', 'Вес_кг', 'Цена', 'Цена_за_кг', 'Цена_за_кг_с_НДС', 'Суммарный_вес', 'Цена_за_кг_итого']
    for col in numeric_cols:
        if col in df.columns:
            df[col] = _normalize_numeric(df[col])

    if 'Дата_статус' in df.columns:
        df['Дата'] = (
            df['Дата_статус']
            .astype(str)
            .str.split(',')
            .str[0]
            .str.strip()
        )
        df['Дата'] = pd.to_datetime(df['Дата'], errors='coerce', dayfirst=True)
        df['Статус'] = (
            df['Дата_статус']
            .astype(str)
            .str.split(',')
            .str[1:]
            .str.join(',')
            .str.strip()
        )

    return df


def _preview_table(df: pd.DataFrame, limit: int = 20) -> PreviewTable:
    sample = df.head(limit).copy()
    sample = sample.fillna('')
    return PreviewTable(columns=list(sample.columns), rows=sample.to_dict(orient='records'))


def _describe_dataframe(df: pd.DataFrame, columns: List[str]) -> Optional[PreviewTable]:
    available = [col for col in columns if col in df.columns]
    if not available:
        return None

    describe_df = df[available].describe().T.reset_index().rename(columns={'index': 'Показатель'})
    rename_map = {
        'count': 'Количество',
        'mean': 'Среднее значение',
        'std': 'Стандартное отклонение',
        'min': 'Минимум',
        '25%': '25-й перцентиль',
        '50%': 'Медиана (50%)',
        '75%': '75-й перцентиль',
        'max': 'Максимум'
    }
    describe_df = describe_df.rename(columns=rename_map)
    describe_df = describe_df.round(2).replace(np.nan, '')
    return _preview_table(describe_df, limit=len(describe_df))


def _generate_tender_metrics(df: pd.DataFrame) -> List[Dict[str, Any]]:
    metrics: List[Dict[str, Any]] = [
        {'label': 'Всего записей', 'value': len(df)}
    ]
    if 'Тип раздела' in df.columns:
        metrics.append({'label': 'Разделов', 'value': df['Тип раздела'].nunique()})
    if 'Заказчик' in df.columns:
        metrics.append({'label': 'Уникальных заказчиков', 'value': df['Заказчик'].nunique()})
    if 'Сумма' in df.columns:
        metrics.append({'label': 'Записей с суммой', 'value': int(df['Сумма'].notna().sum())})
    return metrics


def _generate_product_metrics(df: pd.DataFrame) -> List[Dict[str, Any]]:
    total_weight = float(df['Вес_кг'].sum()) if 'Вес_кг' in df.columns else 0.0
    total_price = float(df['Цена'].sum()) if 'Цена' in df.columns else 0.0
    metrics = [
        {'label': 'Позиций', 'value': len(df)}
    ]
    metrics.append({'label': 'Общий вес (кг)', 'value': f"{total_weight:,.0f}".replace(',', ' ')})
    metrics.append({'label': 'Общая сумма (руб)', 'value': f"{total_price:,.0f}".replace(',', ' ')})
    avg_price = (total_price / total_weight) if total_weight > 0 else 0
    metrics.append({'label': 'Средняя цена за кг', 'value': f"{avg_price:,.1f}".replace(',', ' ')})
    return metrics


def _generate_tender_charts(df: pd.DataFrame) -> List[Dict[str, Any]]:
    charts: List[Dict[str, Any]] = []

    if 'Тип раздела' in df.columns and not df['Тип раздела'].dropna().empty:
        fig, axes = plt.subplots(1, 2, figsize=(14, 6))
        counts = df['Тип раздела'].value_counts()
        axes[0].pie(counts.values, labels=counts.index, autopct='%1.1f%%', startangle=90)
        axes[0].set_title('Распределение по разделам')
        counts.plot(kind='bar', ax=axes[1], color='skyblue')
        axes[1].set_title('Количество записей по разделам')
        axes[1].tick_params(axis='x', rotation=45)
        charts.append({'title': 'Распределение по разделам', 'image': _fig_to_base64(fig)})

    if 'Заказчик' in df.columns and not df['Заказчик'].dropna().empty:
        top_customers = df['Заказчик'].value_counts().head(10)
        fig, ax = plt.subplots(figsize=(10, 6))
        top_customers.sort_values().plot(kind='barh', ax=ax, color='lightgreen')
        ax.set_title('Топ-10 заказчиков по количеству заявок')
        ax.set_xlabel('Количество заявок')
        charts.append({'title': 'Топ заказчиков', 'image': _fig_to_base64(fig)})

    if 'Сумма' in df.columns and df['Сумма'].notna().any():
        fig, ax = plt.subplots(figsize=(10, 6))
        df['Сумма'].plot(kind='hist', bins=20, ax=ax, color='plum', alpha=0.8)
        ax.set_title('Распределение сумм тендеров')
        ax.set_xlabel('Сумма')
        charts.append({'title': 'Распределение сумм', 'image': _fig_to_base64(fig)})

        if 'Тип раздела' in df.columns:
            fig, ax = plt.subplots(figsize=(10, 6))
            section_sums = df.groupby('Тип раздела')['Сумма'].sum().sort_values(ascending=False)
            section_sums.plot(kind='bar', ax=ax, color='coral')
            ax.set_title('Сумма по разделам')
            ax.set_ylabel('Сумма')
            ax.tick_params(axis='x', rotation=45)
            charts.append({'title': 'Сумма по разделам', 'image': _fig_to_base64(fig)})

    return charts


def _generate_product_charts(df: pd.DataFrame) -> List[Dict[str, Any]]:
    charts: List[Dict[str, Any]] = []

    if 'Заказчик' in df.columns and not df['Заказчик'].dropna().empty:
        top_orders = df['Заказчик'].value_counts().head(10)
        fig, ax = plt.subplots(figsize=(10, 6))
        top_orders.plot(kind='bar', ax=ax, color='steelblue')
        ax.set_title('Топ-10 заказчиков по количеству заказов')
        ax.set_ylabel('Количество заказов')
        ax.tick_params(axis='x', rotation=45)
        charts.append({'title': 'Топ заказчиков по количеству', 'image': _fig_to_base64(fig)})

    if {'Заказчик', 'Вес_кг'}.issubset(df.columns) and df['Вес_кг'].notna().any():
        weight_by_customer = df.groupby('Заказчик')['Вес_кг'].sum().sort_values(ascending=False).head(10)
        fig, ax = plt.subplots(figsize=(10, 6))
        weight_by_customer.plot(kind='bar', ax=ax, color='lightgreen')
        ax.set_title('Топ-10 заказчиков по общему весу')
        ax.set_ylabel('Вес (кг)')
        ax.tick_params(axis='x', rotation=45)
        charts.append({'title': 'Вес по заказчикам', 'image': _fig_to_base64(fig)})

    if 'Менеджер' in df.columns and not df['Менеджер'].dropna().empty:
        top_managers = df['Менеджер'].value_counts().head(10)
        fig, ax = plt.subplots(figsize=(10, 6))
        top_managers.plot(kind='bar', ax=ax, color='orange')
        ax.set_title('Топ-10 менеджеров по количеству заказов')
        ax.set_ylabel('Количество заказов')
        ax.tick_params(axis='x', rotation=45)
        charts.append({'title': 'Аналитика по менеджерам', 'image': _fig_to_base64(fig)})

        if {'Цена', 'Количество'}.issubset(df.columns) and df['Цена'].notna().any():
            revenue_by_manager = (
                df.assign(_total=df['Цена'] * df['Количество'].fillna(1))
                .groupby('Менеджер')['_total']
                .sum()
                .sort_values(ascending=False)
                .head(10)
            )
            fig, ax = plt.subplots(figsize=(10, 6))
            revenue_by_manager.plot(kind='bar', ax=ax, color='seagreen')
            ax.set_title('Топ-10 менеджеров по сумме выигранных тендеров')
            ax.set_ylabel('Сумма (руб)')
            ax.tick_params(axis='x', rotation=45)
            charts.append({'title': 'Менеджеры по сумме выигранных тендеров', 'image': _fig_to_base64(fig)})

    if 'Материал' in df.columns and not df['Материал'].dropna().empty:
        material_counts = df['Материал'].value_counts().head(10)
        fig, ax = plt.subplots(figsize=(10, 6))
        ax.pie(material_counts.values, labels=material_counts.index, autopct='%1.1f%%', startangle=90)
        ax.axis('equal')
        ax.set_title('Популярные материалы (в процентах)')
        charts.append({'title': 'Популярные материалы', 'image': _fig_to_base64(fig)})

        if {'Материал', 'Цена_за_кг'}.issubset(df.columns) and df['Цена_за_кг'].notna().any():
            top_materials = material_counts.index.tolist()
            avg_price = (
                df[df['Материал'].isin(top_materials)]
                .groupby('Материал')['Цена_за_кг']
                .mean()
                .reindex(top_materials)
            )
            fig, ax = plt.subplots(figsize=(10, 6))
            avg_price.plot(kind='bar', ax=ax, color='slateblue')
            ax.set_title('Средняя цена за кг по популярным материалам')
            ax.set_ylabel('Цена за кг (руб)')
            ax.tick_params(axis='x', rotation=45)
            if ax.get_legend():
                ax.get_legend().remove()
            for idx, value in enumerate(avg_price.values):
                ax.text(idx, value, f"{value:,.0f}".replace(',', ' '), ha='center', va='bottom', fontsize=10, fontweight='600')
            charts.append({'title': 'Средняя цена популярных материалов', 'image': _fig_to_base64(fig)})

    return charts


def _build_download_payload(df: pd.DataFrame, filename_prefix: str) -> Dict[str, str]:
    csv_data = df.to_csv(index=False, encoding='utf-8-sig')
    csv_b64 = base64.b64encode(csv_data.encode('utf-8-sig')).decode('utf-8')
    excel_buffer = BytesIO()
    df.to_excel(excel_buffer, index=False)
    excel_buffer.seek(0)
    xlsx_b64 = base64.b64encode(excel_buffer.read()).decode('utf-8')
    return {
        'csv': csv_b64,
        'excel': xlsx_b64,
        'csv_name': f'{filename_prefix}_processed.csv',
        'excel_name': f'{filename_prefix}_processed.xlsx'
    }


def analyze_excel(file_storage) -> Dict[str, Any]:
    filename = (file_storage.filename or 'dataset').strip()
    if not filename.lower().endswith(('.xls', '.xlsx')):
        raise AnalyticsProcessingError('Поддерживаются только файлы Excel форматов .xls и .xlsx.')

    file_storage.stream.seek(0)
    binary = file_storage.read()
    if not binary:
        raise AnalyticsProcessingError('Получен пустой файл.')

    buffer = BytesIO(binary)
    try:
        if filename.lower().endswith('.xls'):
            df_raw = pd.read_excel(buffer, header=None, engine='xlrd')
        else:
            df_raw = pd.read_excel(buffer, header=None)
    except Exception as exc:  # pragma: no cover - чтение файла
        raise AnalyticsProcessingError(f'Не удалось прочитать файл: {exc}') from exc

    if df_raw.empty:
        raise AnalyticsProcessingError('Файл не содержит данных.')

    data_type = detect_data_type(df_raw)
    if data_type == 'tender':
        processed = process_tender_data(df_raw)
        if processed is None or processed.empty:
            raise AnalyticsProcessingError('Не удалось извлечь данные по тендерам. Проверьте структуру файла.')
        metrics = _generate_tender_metrics(processed)
        charts = _generate_tender_charts(processed)
        stats = _describe_dataframe(processed, ['Сумма', 'Количество'])
        analysis_label = 'Данные по тендерам'
    else:
        processed = process_product_data(df_raw)
        if processed.empty:
            raise AnalyticsProcessingError('Не удалось извлечь данные по продукции. Проверьте структуру файла.')
        metrics = _generate_product_metrics(processed)
        charts = _generate_product_charts(processed)
        stats = _describe_dataframe(processed, ['Количество', 'Вес_кг', 'Цена', 'Цена_за_кг'])
        analysis_label = 'Данные по продукции'

    preview = _preview_table(processed)
    download_payload = _build_download_payload(processed, filename_prefix=filename.rsplit('.', 1)[0])

    return {
        'analysis_type': data_type,
        'analysis_label': analysis_label,
        'metrics': metrics,
        'charts': charts,
        'preview': preview,
        'stats': stats,
        'download': download_payload,
        'columns': list(processed.columns)
    }
