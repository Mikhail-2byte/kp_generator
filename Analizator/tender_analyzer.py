import streamlit as st
import pandas as pd
import matplotlib.pyplot as plt
import seaborn as sns
import numpy as np
from datetime import datetime
import warnings
warnings.filterwarnings('ignore')

# Настройки
st.set_page_config(page_title="Универсальный анализатор таблиц", layout="wide")
st.title("📊 Универсальный анализатор Excel таблиц")

def detect_data_type(df_raw):
    """Автоматическое определение типа данных"""
    # Анализ структуры данных
    sample_text = df_raw.astype(str).iloc[:10].to_string()
    
    # Ключевые слова для определения типа данных
    tender_keywords = ['Мониторинг', 'участвовали', 'Подались', 'Проиграли', 'Отменен', 'тендер', 'Тендер']
    product_keywords = ['ч.', 'заказчик', 'вес', 'цена', 'количество', 'материал']
    
    tender_score = sum(1 for keyword in tender_keywords if keyword.lower() in sample_text.lower())
    product_score = sum(1 for keyword in product_keywords if keyword.lower() in sample_text.lower())
    
    if tender_score > product_score:
        return "tender"
    else:
        return "product"

def process_tender_data(df_raw):
    """Обработка данных по тендерам"""
    # Поиск заголовков разделов
    sections = []
    current_section = None
    
    for idx, row in df_raw.iterrows():
        if pd.notna(row[0]) and any(keyword in str(row[0]) for keyword in 
                                  ['Мониторинг', 'участвовали', 'Отменен', 'Подались', 'Проиграли', 'Итого']):
            current_section = row[0]
            sections.append((current_section, idx))
    
    # Обработка данных для каждого раздела
    all_data = []
    
    for i, (section_name, start_idx) in enumerate(sections):
        end_idx = sections[i + 1][1] if i + 1 < len(sections) else len(df_raw)
        section_data = df_raw.iloc[start_idx:end_idx].copy()
        
        # Находим строку с заголовками
        header_row = None
        for j in range(len(section_data)):
            if pd.notna(section_data.iloc[j, 0]) and str(section_data.iloc[j, 0]).isdigit():
                header_row = j
                break
        
        if header_row is not None:
            headers = []
            for col in range(len(section_data.columns)):
                header_val = section_data.iloc[header_row-1, col] if header_row > 0 else f"Column_{col}"
                if pd.isna(header_val) or header_val == '':
                    headers.append(f"Column_{col}")
                else:
                    headers.append(str(header_val))
            
            data_part = section_data.iloc[header_row:].copy()
            data_part.columns = headers
            data_part['Раздел'] = section_name
            
            # Очищаем пустые строки
            data_part = data_part.dropna(subset=[headers[0]])
            data_part = data_part[data_part[headers[0]].astype(str).str.isdigit()]
            
            if not data_part.empty:
                all_data.append(data_part)
    
    if all_data:
        df = pd.concat(all_data, ignore_index=True)
        
        # Стандартизируем имена колонок
        column_mapping = {
            'A': '№', 'B': 'Дата начала', 'C': 'Дата окончания', 'D': 'Номер заявки',
            'E': 'Заказчик', 'F': 'Тендер', 'G': 'Признак', 'H': 'Тип раздела',
            'I': 'Площадка', 'J': 'Комментарий', 'K': 'Сумма', 'L': 'Количество'
        }
        
        current_columns = df.columns.tolist()
        new_columns = []
        for i, col in enumerate(current_columns):
            if i < len(column_mapping):
                new_columns.append(list(column_mapping.values())[i])
            else:
                new_columns.append(f'Column_{i}')
        
        df.columns = new_columns
        df = df.replace('', np.nan)
        
        # Преобразование данных
        date_columns = ['Дата начала', 'Дата окончания']
        for col in date_columns:
            if col in df.columns:
                df[col] = pd.to_datetime(df[col], errors='coerce', dayfirst=True)
        
        if 'Сумма' in df.columns:
            df['Сумма'] = pd.to_numeric(df['Сумма'], errors='coerce')
        
        return df
    return None

def process_product_data(df_raw):
    """Обработка данных по продукции"""
    # Создаем копию данных
    df = df_raw.copy()
    
    # Определяем заголовки на основе анализа структуры
    headers = [
        'Наименование', 'Чертеж', 'Заказчик', 'Количество', 'Вес_кг', 'Материал',
        'Цена', 'Цена_за_кг', 'Цена_за_кг_с_НДС', 'Менеджер', 'Дата_статус',
        'Тендер', 'Площадка', 'Суммарный_вес', 'Цена_за_кг_итого'
    ]
    
    # Если количество столбцов не совпадает, корректируем
    if len(df.columns) > len(headers):
        df = df.iloc[:, :len(headers)]
    elif len(df.columns) < len(headers):
        for i in range(len(df.columns), len(headers)):
            df[i] = np.nan
    
    df.columns = headers
    
    # Очистка данных - удаляем строки с группировками (где нет номера чертежа)
    df = df[df['Чертеж'].notna() & (df['Чертеж'] != '')]
    
    # Преобразование числовых колонок
    numeric_cols = ['Количество', 'Вес_кг', 'Цена', 'Цена_за_кг', 'Цена_за_кг_с_НДС', 'Суммарный_вес', 'Цена_за_кг_итого']
    for col in numeric_cols:
        if col in df.columns:
            df[col] = pd.to_numeric(df[col], errors='coerce')
    
    # Обработка дат
    if 'Дата_статус' in df.columns:
        df['Дата'] = df['Дата_статус'].str.split(',').str[0]
        df['Дата'] = pd.to_datetime(df['Дата'], errors='coerce', dayfirst=True)
        df['Статус'] = df['Дата_статус'].str.split(',').str[1:].str.join(',').str.strip()
    
    return df

# Загрузка файла
uploaded_file = st.file_uploader("Загрузите Excel файл", type=['xlsx', 'xls'])

if uploaded_file is not None:
    try:
        # Чтение файла
        df_raw = pd.read_excel(uploaded_file, header=None)
        
        # Автоматическое определение типа данных
        data_type = detect_data_type(df_raw)
        
        if data_type == "tender":
            st.success("🔍 Обнаружены данные по тендерам")
            df = process_tender_data(df_raw)
            analysis_type = "tender"
        else:
            st.success("🏭 Обнаружены данные по продукции")
            df = process_product_data(df_raw)
            analysis_type = "product"
        
        if df is not None and not df.empty:
            # Основная информация
            st.subheader("📋 Обзор данных")
            st.write(f"**Тип данных:** {'Тендеры' if analysis_type == 'tender' else 'Продукция'}")
            st.write(f"**Размер:** {df.shape[0]} строк, {df.shape[1]} столбцов")
            
            col1, col2, col3, col4 = st.columns(4)
            
            if analysis_type == "tender":
                with col1:
                    st.metric("Всего записей", len(df))
                with col2:
                    st.metric("Разделов", df['Тип раздела'].nunique() if 'Тип раздела' in df.columns else 0)
                with col3:
                    st.metric("Заказчиков", df['Заказчик'].nunique() if 'Заказчик' in df.columns else 0)
                with col4:
                    has_sum = 'Сумма' in df.columns and df['Сумма'].notna().any()
                    st.metric("Записей с суммами", df['Сумма'].notna().sum() if has_sum else 0)
            else:
                with col1:
                    st.metric("Позиций", len(df))
                with col2:
                    total_weight = df['Вес_кг'].sum() if 'Вес_кг' in df.columns else 0
                    st.metric("Общий вес (кг)", f"{total_weight:,.0f}")
                with col3:
                    total_price = df['Цена'].sum() if 'Цена' in df.columns else 0
                    st.metric("Общая сумма", f"{total_price:,.0f} руб")
                with col4:
                    avg_price_per_kg = total_price / total_weight if total_weight > 0 else 0
                    st.metric("Ср. цена за кг", f"{avg_price_per_kg:.1f} руб")
            
            # Просмотр данных
            st.subheader("🔍 Просмотр данных")
            st.dataframe(df.head(20), use_container_width=True)
            
            # Аналитика в зависимости от типа данных
            st.subheader("📈 Аналитика и визуализация")
            
            if analysis_type == "tender":
                # Аналитика для тендеров
                analysis_option = st.selectbox(
                    "Выберите тип анализа:",
                    ["Распределение по разделам", "Топ заказчиков", "Динамика по датам", "Анализ сумм"]
                )
                
                if analysis_option == "Распределение по разделам" and 'Тип раздела' in df.columns:
                    fig, (ax1, ax2) = plt.subplots(1, 2, figsize=(15, 6))
                    
                    section_counts = df['Тип раздела'].value_counts()
                    ax1.pie(section_counts.values, labels=section_counts.index, autopct='%1.1f%%', startangle=90)
                    ax1.set_title('Распределение по разделам')
                    
                    section_counts.plot(kind='bar', ax=ax2, color='skyblue')
                    ax2.set_title('Количество записей по разделам')
                    ax2.set_ylabel('Количество')
                    ax2.tick_params(axis='x', rotation=45)
                    
                    plt.tight_layout()
                    st.pyplot(fig)
                
                elif analysis_option == "Топ заказчиков" and 'Заказчик' in df.columns:
                    top_n = st.slider("Количество топ заказчиков:", 5, 20, 10)
                    customer_counts = df['Заказчик'].value_counts().head(top_n)
                    
                    fig, ax = plt.subplots(figsize=(12, 6))
                    customer_counts.plot(kind='barh', ax=ax, color='lightgreen')
                    ax.set_title(f'Топ-{top_n} заказчиков по количеству заявок')
                    ax.set_xlabel('Количество заявок')
                    
                    for i, v in enumerate(customer_counts.values):
                        ax.text(v + 0.1, i, str(v), va='center')
                    
                    plt.tight_layout()
                    st.pyplot(fig)
                
                elif analysis_option == "Анализ сумм" and 'Сумма' in df.columns and df['Сумма'].notna().any():
                    col1, col2 = st.columns(2)
                    
                    with col1:
                        fig1, ax1 = plt.subplots(figsize=(10, 6))
                        df['Сумма'].hist(bins=20, ax=ax1, alpha=0.7, color='purple')
                        ax1.set_title('Распределение сумм тендеров')
                        ax1.set_xlabel('Сумма')
                        ax1.set_ylabel('Количество')
                        st.pyplot(fig1)
                    
                    with col2:
                        if 'Тип раздела' in df.columns:
                            fig2, ax2 = plt.subplots(figsize=(10, 6))
                            section_sums = df.groupby('Тип раздела')['Сумма'].sum().sort_values(ascending=False)
                            section_sums.plot(kind='bar', ax=ax2, color='coral')
                            ax2.set_title('Общая сумма по разделам')
                            ax2.set_ylabel('Сумма')
                            ax2.tick_params(axis='x', rotation=45)
                            plt.tight_layout()
                            st.pyplot(fig2)
            
            else:
                # Аналитика для продукции
                analysis_option = st.selectbox(
                    "Выберите тип анализа:",
                    ["Топ заказчиков", "Анализ по менеджерам", "Распределение по весу", "Ценовой анализ", "Анализ материалов"]
                )
                
                if analysis_option == "Топ заказчиков" and 'Заказчик' in df.columns:
                    col1, col2 = st.columns(2)
                    
                    with col1:
                        # По количеству заказов
                        customer_orders = df['Заказчик'].value_counts().head(10)
                        fig, ax = plt.subplots(figsize=(10, 6))
                        customer_orders.plot(kind='bar', ax=ax, color='lightblue')
                        ax.set_title('Топ-10 заказчиков по количеству заказов')
                        ax.set_ylabel('Количество заказов')
                        ax.tick_params(axis='x', rotation=45)
                        st.pyplot(fig)
                    
                    with col2:
                        # По общему весу
                        if 'Вес_кг' in df.columns:
                            customer_weight = df.groupby('Заказчик')['Вес_кг'].sum().sort_values(ascending=False).head(10)
                            fig, ax = plt.subplots(figsize=(10, 6))
                            customer_weight.plot(kind='bar', ax=ax, color='lightgreen')
                            ax.set_title('Топ-10 заказчиков по общему весу')
                            ax.set_ylabel('Вес (кг)')
                            ax.tick_params(axis='x', rotation=45)
                            st.pyplot(fig)
                
                elif analysis_option == "Анализ по менеджерам" and 'Менеджер' in df.columns:
                    col1, col2 = st.columns(2)
                    
                    with col1:
                        manager_orders = df['Менеджер'].value_counts().head(10)
                        fig, ax = plt.subplots(figsize=(10, 6))
                        manager_orders.plot(kind='bar', ax=ax, color='orange')
                        ax.set_title('Топ-10 менеджеров по количеству заказов')
                        ax.set_ylabel('Количество заказов')
                        ax.tick_params(axis='x', rotation=45)
                        st.pyplot(fig)
                    
                    with col2:
                        if 'Цена' in df.columns:
                            manager_revenue = df.groupby('Менеджер')['Цена'].sum().sort_values(ascending=False).head(10)
                            fig, ax = plt.subplots(figsize=(10, 6))
                            manager_revenue.plot(kind='bar', ax=ax, color='red')
                            ax.set_title('Топ-10 менеджеров по выручке')
                            ax.set_ylabel('Выручка (руб)')
                            ax.tick_params(axis='x', rotation=45)
                            st.pyplot(fig)
                
                elif analysis_option == "Распределение по весу" and 'Вес_кг' in df.columns:
                    col1, col2 = st.columns(2)
                    
                    with col1:
                        fig, ax = plt.subplots(figsize=(10, 6))
                        df['Вес_кг'].hist(bins=30, ax=ax, alpha=0.7, color='purple')
                        ax.set_title('Распределение веса заказов')
                        ax.set_xlabel('Вес (кг)')
                        ax.set_ylabel('Количество')
                        st.pyplot(fig)
                    
                    with col2:
                        # Боксплот для выбросов
                        fig, ax = plt.subplots(figsize=(10, 6))
                        df['Вес_кг'].plot(kind='box', ax=ax)
                        ax.set_title('Боксплот веса заказов')
                        ax.set_ylabel('Вес (кг)')
                        st.pyplot(fig)
                
                elif analysis_option == "Ценовой анализ" and 'Цена_за_кг' in df.columns:
                    col1, col2 = st.columns(2)
                    
                    with col1:
                        fig, ax = plt.subplots(figsize=(10, 6))
                        df['Цена_за_кг'].hist(bins=30, ax=ax, alpha=0.7, color='green')
                        ax.set_title('Распределение цены за кг')
                        ax.set_xlabel('Цена за кг (руб)')
                        ax.set_ylabel('Количество')
                        st.pyplot(fig)
                    
                    with col2:
                        if 'Материал' in df.columns:
                            # Средняя цена по материалам
                            material_prices = df.groupby('Материал')['Цена_за_кг'].mean().sort_values(ascending=False).head(10)
                            fig, ax = plt.subplots(figsize=(10, 6))
                            material_prices.plot(kind='bar', ax=ax, color='brown')
                            ax.set_title('Топ-10 материалов по средней цене за кг')
                            ax.set_ylabel('Средняя цена за кг (руб)')
                            ax.tick_params(axis='x', rotation=45)
                            st.pyplot(fig)
                
                elif analysis_option == "Анализ материалов" and 'Материал' in df.columns:
                    material_counts = df['Материал'].value_counts().head(15)
                    
                    fig, ax = plt.subplots(figsize=(12, 8))
                    material_counts.plot(kind='pie', ax=ax, autopct='%1.1f%%')
                    ax.set_title('Распределение по материалам (топ-15)')
                    ax.set_ylabel('')
                    st.pyplot(fig)
            
            # Детальная статистика
            st.subheader("📊 Детальная статистика")
            
            if analysis_type == "product":
                col1, col2 = st.columns(2)
                
                with col1:
                    st.write("**Статистика по весу:**")
                    if 'Вес_кг' in df.columns:
                        st.write(df['Вес_кг'].describe())
                
                with col2:
                    st.write("**Статистика по цене за кг:**")
                    if 'Цена_за_кг' in df.columns:
                        st.write(df['Цена_за_кг'].describe())
            
            # Экспорт результатов
            st.subheader("💾 Экспорт данных")
            
            if st.button("Скачать обработанные данные (CSV)"):
                csv = df.to_csv(index=False, encoding='utf-8-sig')
                st.download_button(
                    label="📥 Нажмите для скачивания CSV",
                    data=csv,
                    file_name="обработанные_данные.csv",
                    mime="text/csv"
                )
        
        else:
            st.error("Не удалось обработать данные. Проверьте структуру файла.")
    
    except Exception as e:
        st.error(f"Ошибка при обработке файла: {str(e)}")
        st.info("""
        **Рекомендации:**
        - Убедитесь, что файл имеет правильную структуру
        - Для тендерных данных ожидаются разделы: 'Мониторинг цен', 'Не участвовали' и т.д.
        - Для данных по продукции ожидаются колонки с наименованиями, чертежами, заказчиками
        """)

else:
    st.info("👆 Загрузите Excel файл для начала анализа")
    
    st.subheader("Поддерживаемые форматы:")
    col1, col2 = st.columns(2)
    
    with col1:
        st.write("**Данные по тендерам:**")
        st.write("""
        - Разделы: Мониторинг цен, Не участвовали, Подались, Проиграли, Отменен
        - Содержит информацию о заказчиках, датах, суммах
        """)
    
    with col2:
        st.write("**Данные по продукции:**")
        st.write("""
        - Детальная информация о изделиях
        - Чертежи, заказчики, вес, цены
        - Материалы, менеджеры
        """)