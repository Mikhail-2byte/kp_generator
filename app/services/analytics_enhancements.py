"""Расширенные функции аналитики: графики маржинальности, динамика курсов, интерактивные отчеты."""

from datetime import datetime, timedelta
from typing import Any, Dict, List, Optional

import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt
import numpy as np
import pandas as pd
import seaborn as sns

from app.services.repositories import generation_repository

sns.set_theme(style='whitegrid')


def generate_margin_analysis(user_id: Optional[int] = None, days: int = 30) -> Dict[str, Any]:
    """
    Генерирует анализ маржинальности за указанный период.
    
    Args:
        user_id: ID пользователя (если None, анализирует всех)
        days: Количество дней для анализа
    
    Returns:
        Словарь с графиками и метриками маржинальности
    """
    from flask import current_app
    
    date_from = (datetime.now() - timedelta(days=days)).strftime('%Y-%m-%d')
    app_config = current_app.config['APP_SETTINGS']
    
    history = generation_repository.get_history(
        app_config,
        page=1,
        per_page=1000,
        date_from=date_from
    )
    
    if not history['items']:
        return {
            'charts': [],
            'metrics': [],
            'message': 'Нет данных за указанный период'
        }
    
    # Подготавливаем данные
    margins = []
    revenues = []
    costs = []
    dates = []
    
    for item in history['items']:
        margin = item.get('margin_percent', 0)
        revenue = item.get('total_general_price', item.get('final_price', 0) * item.get('quantity', 0))
        cost = revenue / (1 + margin / 100) if margin > 0 else revenue
        
        margins.append(margin)
        revenues.append(revenue)
        costs.append(cost)
        dates.append(item.get('timestamp', ''))
    
    charts = []
    metrics = []
    
    # График динамики маржи
    if margins:
        fig, ax = plt.subplots(figsize=(12, 6))
        ax.plot(range(len(margins)), margins, marker='o', linewidth=2, markersize=4)
        ax.set_title('Динамика маржинальности', fontsize=14, fontweight='bold')
        ax.set_xlabel('Номер генерации')
        ax.set_ylabel('Маржа (%)')
        ax.grid(True, alpha=0.3)
        ax.axhline(y=np.mean(margins), color='r', linestyle='--', label=f'Средняя маржа: {np.mean(margins):.1f}%')
        ax.legend()
        plt.tight_layout()
        
        from io import BytesIO
        import base64
        buffer = BytesIO()
        plt.savefig(buffer, format='png', bbox_inches='tight')
        buffer.seek(0)
        chart_b64 = base64.b64encode(buffer.read()).decode('utf-8')
        plt.close(fig)
        
        charts.append({
            'title': 'Динамика маржинальности',
            'image': chart_b64,
            'type': 'line'
        })
        
        # Метрики
        metrics.extend([
            {'label': 'Средняя маржа', 'value': f"{np.mean(margins):.2f}%"},
            {'label': 'Медианная маржа', 'value': f"{np.median(margins):.2f}%"},
            {'label': 'Минимальная маржа', 'value': f"{np.min(margins):.2f}%"},
            {'label': 'Максимальная маржа', 'value': f"{np.max(margins):.2f}%"},
        ])
    
    # Распределение маржи
    if margins:
        fig, ax = plt.subplots(figsize=(10, 6))
        ax.hist(margins, bins=20, edgecolor='black', alpha=0.7)
        ax.set_title('Распределение маржинальности', fontsize=14, fontweight='bold')
        ax.set_xlabel('Маржа (%)')
        ax.set_ylabel('Количество генераций')
        ax.axvline(x=np.mean(margins), color='r', linestyle='--', label=f'Средняя: {np.mean(margins):.1f}%')
        ax.legend()
        plt.tight_layout()
        
        from io import BytesIO
        import base64
        buffer = BytesIO()
        plt.savefig(buffer, format='png', bbox_inches='tight')
        buffer.seek(0)
        chart_b64 = base64.b64encode(buffer.read()).decode('utf-8')
        plt.close(fig)
        
        charts.append({
            'title': 'Распределение маржинальности',
            'image': chart_b64,
            'type': 'histogram'
        })
    
    # График маржа vs выручка
    if margins and revenues:
        fig, ax = plt.subplots(figsize=(10, 6))
        scatter = ax.scatter(revenues, margins, alpha=0.6, s=50)
        ax.set_title('Маржа vs Выручка', fontsize=14, fontweight='bold')
        ax.set_xlabel('Выручка (руб)')
        ax.set_ylabel('Маржа (%)')
        ax.grid(True, alpha=0.3)
        plt.tight_layout()
        
        from io import BytesIO
        import base64
        buffer = BytesIO()
        plt.savefig(buffer, format='png', bbox_inches='tight')
        buffer.seek(0)
        chart_b64 = base64.b64encode(buffer.read()).decode('utf-8')
        plt.close(fig)
        
        charts.append({
            'title': 'Маржа vs Выручка',
            'image': chart_b64,
            'type': 'scatter'
        })
    
    return {
        'charts': charts,
        'metrics': metrics,
        'period_days': days
    }


def generate_exchange_rate_analysis(days: int = 90) -> Dict[str, Any]:
    """
    Генерирует анализ динамики курсов валют (если данные доступны).
    
    Args:
        days: Количество дней для анализа
    
    Returns:
        Словарь с графиками динамики курсов
    """
    # Заглушка для будущей реализации с реальными данными курсов
    # Пока возвращаем структуру для интеграции
    
    charts = []
    metrics = []
    
    # Пример графика (в реальности данные будут из внешнего API или БД)
    fig, ax = plt.subplots(figsize=(12, 6))
    
    # Генерируем примерные данные
    dates = pd.date_range(end=datetime.now(), periods=days, freq='D')
    rates = 12 + np.random.normal(0, 0.5, days)  # Примерные курсы юаня
    
    ax.plot(dates, rates, linewidth=2)
    ax.set_title('Динамика курса юаня к рублю', fontsize=14, fontweight='bold')
    ax.set_xlabel('Дата')
    ax.set_ylabel('Курс (руб/юань)')
    ax.grid(True, alpha=0.3)
    plt.xticks(rotation=45)
    plt.tight_layout()
    
    from io import BytesIO
    import base64
    buffer = BytesIO()
    plt.savefig(buffer, format='png', bbox_inches='tight')
    buffer.seek(0)
    chart_b64 = base64.b64encode(buffer.read()).decode('utf-8')
    plt.close(fig)
    
    charts.append({
        'title': 'Динамика курса юаня',
        'image': chart_b64,
        'type': 'line'
    })
    
    metrics.extend([
        {'label': 'Средний курс', 'value': f"{np.mean(rates):.2f} руб/юань"},
        {'label': 'Минимальный курс', 'value': f"{np.min(rates):.2f} руб/юань"},
        {'label': 'Максимальный курс', 'value': f"{np.max(rates):.2f} руб/юань"},
    ])
    
    return {
        'charts': charts,
        'metrics': metrics,
        'note': 'Данные курсов являются примерными. Для реальных данных требуется интеграция с внешним API.'
    }


def generate_interactive_report(
    user_id: Optional[int] = None,
    date_from: Optional[str] = None,
    date_to: Optional[str] = None
) -> Dict[str, Any]:
    """
    Генерирует интерактивный отчет с различными метриками и графиками.
    
    Args:
        user_id: ID пользователя (если None, анализирует всех)
        date_from: Начальная дата (формат: YYYY-MM-DD)
        date_to: Конечная дата (формат: YYYY-MM-DD)
    
    Returns:
        Словарь с комплексным отчетом
    """
    from flask import current_app
    
    app_config = current_app.config['APP_SETTINGS']
    
    history = generation_repository.get_history(
        app_config,
        page=1,
        per_page=1000,
        date_from=date_from,
        date_to=date_to
    )
    
    if not history['items']:
        return {
            'summary': {},
            'charts': [],
            'tables': [],
            'message': 'Нет данных за указанный период'
        }
    
    # Подготавливаем данные
    df_data = []
    for item in history['items']:
        df_data.append({
            'date': item.get('timestamp', ''),
            'company': item.get('company', ''),
            'margin': item.get('margin_percent', 0),
            'revenue': item.get('total_general_price', item.get('final_price', 0) * item.get('quantity', 0)),
            'quantity': item.get('quantity', 0),
            'positions_count': item.get('positions_count', 1),
        })
    
    df = pd.DataFrame(df_data)
    
    # Сводная статистика
    summary = {
        'total_generations': len(df),
        'total_revenue': float(df['revenue'].sum()),
        'avg_margin': float(df['margin'].mean()),
        'total_quantity': int(df['quantity'].sum()),
        'unique_companies': df['company'].nunique(),
    }
    
    charts = []
    
    # График выручки по дням
    if 'date' in df.columns and not df['date'].empty:
        df['date'] = pd.to_datetime(df['date'], errors='coerce')
        daily_revenue = df.groupby(df['date'].dt.date)['revenue'].sum()
        
        fig, ax = plt.subplots(figsize=(12, 6))
        daily_revenue.plot(kind='bar', ax=ax, color='steelblue')
        ax.set_title('Выручка по дням', fontsize=14, fontweight='bold')
        ax.set_xlabel('Дата')
        ax.set_ylabel('Выручка (руб)')
        ax.tick_params(axis='x', rotation=45)
        plt.tight_layout()
        
        from io import BytesIO
        import base64
        buffer = BytesIO()
        plt.savefig(buffer, format='png', bbox_inches='tight')
        buffer.seek(0)
        chart_b64 = base64.b64encode(buffer.read()).decode('utf-8')
        plt.close(fig)
        
        charts.append({
            'title': 'Выручка по дням',
            'image': chart_b64,
            'type': 'bar'
        })
    
    # Топ компаний по выручке
    if 'company' in df.columns:
        top_companies = df.groupby('company')['revenue'].sum().sort_values(ascending=False).head(10)
        
        fig, ax = plt.subplots(figsize=(10, 6))
        top_companies.plot(kind='barh', ax=ax, color='lightgreen')
        ax.set_title('Топ-10 компаний по выручке', fontsize=14, fontweight='bold')
        ax.set_xlabel('Выручка (руб)')
        plt.tight_layout()
        
        from io import BytesIO
        import base64
        buffer = BytesIO()
        plt.savefig(buffer, format='png', bbox_inches='tight')
        buffer.seek(0)
        chart_b64 = base64.b64encode(buffer.read()).decode('utf-8')
        plt.close(fig)
        
        charts.append({
            'title': 'Топ компаний по выручке',
            'image': chart_b64,
            'type': 'barh'
        })
    
    # Таблицы
    tables = []
    
    # Таблица по компаниям
    if 'company' in df.columns:
        company_stats = df.groupby('company').agg({
            'revenue': 'sum',
            'margin': 'mean',
            'quantity': 'sum'
        }).round(2)
        company_stats.columns = ['Общая выручка', 'Средняя маржа (%)', 'Общее количество']
        company_stats = company_stats.sort_values('Общая выручка', ascending=False).head(20)
        
        tables.append({
            'title': 'Статистика по компаниям',
            'data': company_stats.to_dict(orient='records'),
            'columns': list(company_stats.columns)
        })
    
    return {
        'summary': summary,
        'charts': charts,
        'tables': tables,
        'period': {
            'from': date_from,
            'to': date_to
        }
    }

