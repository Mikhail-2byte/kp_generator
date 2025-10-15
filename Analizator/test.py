import streamlit as st
import pandas as pd
import matplotlib.pyplot as plt
import seaborn as sns
import numpy as np
from datetime import datetime
import warnings
warnings.filterwarnings('ignore')

# Настройки
st.set_page_config(page_title="Анализатор тендеров", layout="wide")
st.title("📊 Анализатор тендерных данных")

# Загрузка файла
uploaded_file = st.file_uploader("Загрузите Excel файл с тендерами", type=['xlsx', 'xls'])

if uploaded_file is not None:
    try:
        # Чтение файла
        df_raw = pd.read_excel(uploaded_file, header=None)
        
        # Автоматическое определение структуры данных
        st.subheader("📋 Структура данных")
        
        # Поиск заголовков разделов
        sections = []
        current_section = None
        data_rows = []
        
        for idx, row in df_raw.iterrows():
            # Поиск названий разделов (обычно в столбце A)
            if pd.notna(row[0]) and any(keyword in str(row[0]) for keyword in 
                                      ['Мониторинг', 'участвовали', 'Отменен', 'Подались', 'Проиграли', 'Итого']):
                current_section = row[0]
                sections.append((current_section, idx))
        
        # Обработка данных для каждого раздела
        all_data = []
        
        for i, (section_name, start_idx) in enumerate(sections):
            # Определяем конец раздела
            end_idx = sections[i + 1][1] if i + 1 < len(sections) else len(df_raw)
            
            # Извлекаем данные раздела
            section_data = df_raw.iloc[start_idx:end_idx].copy()
            
            # Находим строку с заголовками (первая строка с числовыми данными)
            header_row = None
            for j in range(len(section_data)):
                if pd.notna(section_data.iloc[j, 0]) and str(section_data.iloc[j, 0]).isdigit():
                    header_row = j
                    break
            
            if header_row is not None:
                # Используем предыдущую строку как заголовки или создаем стандартные
                headers = []
                for col in range(len(section_data.columns)):
                    header_val = section_data.iloc[header_row-1, col] if header_row > 0 else f"Column_{col}"
                    if pd.isna(header_val) or header_val == '':
                        headers.append(f"Column_{col}")
                    else:
                        headers.append(str(header_val))
                
                # Создаем DataFrame для раздела
                data_part = section_data.iloc[header_row:].copy()
                data_part.columns = headers
                data_part['Раздел'] = section_name
                
                # Очищаем пустые строки
                data_part = data_part.dropna(subset=[headers[0]])
                data_part = data_part[data_part[headers[0]].astype(str).str.isdigit()]
                
                if not data_part.empty:
                    all_data.append(data_part)
        
        # Объединяем все данные
        if all_data:
            df = pd.concat(all_data, ignore_index=True)
            
            # Стандартизируем имена колонок на основе анализа вашего файла
            column_mapping = {
                'A': '№',
                'B': 'Дата начала', 
                'C': 'Дата окончания',
                'D': 'Номер заявки',
                'E': 'Заказчик',
                'F': 'Тендер',
                'G': 'Признак',
                'H': 'Тип раздела',
                'I': 'Площадка',
                'J': 'Комментарий',
                'K': 'Сумма',
                'L': 'Количество'
            }
            
            # Переименовываем колонки
            current_columns = df.columns.tolist()
            new_columns = []
            for i, col in enumerate(current_columns):
                if i < len(column_mapping):
                    new_columns.append(list(column_mapping.values())[i])
                else:
                    new_columns.append(f'Column_{i}')
            
            df.columns = new_columns
            
            # Очистка данных
            df = df.replace('', np.nan)
            
            # Преобразование дат
            date_columns = ['Дата начала', 'Дата окончания']
            for col in date_columns:
                if col in df.columns:
                    df[col] = pd.to_datetime(df[col], errors='coerce', dayfirst=True)
            
            # Преобразование числовых колонок
            if 'Сумма' in df.columns:
                df['Сумма'] = pd.to_numeric(df['Сумма'], errors='coerce')
            
            if 'Количество' in df.columns:
                df['Количество'] = pd.to_numeric(df['Количество'], errors='coerce')
            
            # Основная информация
            st.success(f"✅ Данные успешно загружены: {len(df)} записей")
            
            col1, col2, col3, col4 = st.columns(4)
            with col1:
                st.metric("Всего записей", len(df))
            with col2:
                st.metric("Количество разделов", df['Тип раздела'].nunique())
            with col3:
                st.metric("Уникальных заказчиков", df['Заказчик'].nunique())
            with col4:
                st.metric("Записей с суммами", df['Сумма'].notna().sum())
            
            # Просмотр данных
            st.subheader("🔍 Просмотр данных")
            
            tab1, tab2, tab3 = st.tabs(["Все данные", "По разделам", "Статистика"])
            
            with tab1:
                st.dataframe(df, use_container_width=True)
            
            with tab2:
                selected_section = st.selectbox("Выберите раздел:", df['Тип раздела'].unique())
                section_data = df[df['Тип раздела'] == selected_section]
                st.write(f"**Данные раздела: {selected_section}**")
                st.dataframe(section_data, use_container_width=True)
            
            with tab3:
                st.write("**Основная статистика:**")
                col1, col2 = st.columns(2)
                
                with col1:
                    st.write("**Распределение по разделам:**")
                    section_counts = df['Тип раздела'].value_counts()
                    st.dataframe(section_counts)
                
                with col2:
                    if 'Сумма' in df.columns and df['Сумма'].notna().any():
                        st.write("**Статистика по суммам:**")
                        st.write(df['Сумма'].describe())
            
            # Аналитика и графики
            st.subheader("📈 Аналитика и визуализация")
            
            analysis_type = st.selectbox(
                "Выберите тип анализа:",
                ["Распределение по разделам", "Топ заказчиков", "Динамика по датам", "Анализ сумм"]
            )
            
            if analysis_type == "Распределение по разделам":
                fig, (ax1, ax2) = plt.subplots(1, 2, figsize=(15, 6))
                
                # Круговая диаграмма
                section_counts = df['Тип раздела'].value_counts()
                ax1.pie(section_counts.values, labels=section_counts.index, autopct='%1.1f%%', startangle=90)
                ax1.set_title('Распределение по разделам')
                
                # Столбчатая диаграмма
                section_counts.plot(kind='bar', ax=ax2, color='skyblue')
                ax2.set_title('Количество записей по разделам')
                ax2.set_ylabel('Количество')
                ax2.tick_params(axis='x', rotation=45)
                
                plt.tight_layout()
                st.pyplot(fig)
            
            elif analysis_type == "Топ заказчиков":
                top_n = st.slider("Количество топ заказчиков:", 5, 20, 10)
                
                customer_counts = df['Заказчик'].value_counts().head(top_n)
                
                fig, ax = plt.subplots(figsize=(12, 6))
                customer_counts.plot(kind='barh', ax=ax, color='lightgreen')
                ax.set_title(f'Топ-{top_n} заказчиков по количеству заявок')
                ax.set_xlabel('Количество заявок')
                
                # Добавляем подписи значений
                for i, v in enumerate(customer_counts.values):
                    ax.text(v + 0.1, i, str(v), va='center')
                
                plt.tight_layout()
                st.pyplot(fig)
            
            elif analysis_type == "Динамика по датам" and 'Дата начала' in df.columns:
                # Группировка по месяцам
                df_date = df.copy()
                df_date['Месяц'] = df_date['Дата начала'].dt.to_period('M')
                monthly_counts = df_date.groupby('Месяц').size()
                
                fig, ax = plt.subplots(figsize=(12, 6))
                monthly_counts.plot(kind='line', marker='o', ax=ax, color='orange')
                ax.set_title('Динамика количества заявок по месяцам')
                ax.set_ylabel('Количество заявок')
                ax.set_xlabel('Месяц')
                ax.grid(True, alpha=0.3)
                
                plt.tight_layout()
                st.pyplot(fig)
            
            elif analysis_type == "Анализ сумм" and 'Сумма' in df.columns and df['Сумма'].notna().any():
                col1, col2 = st.columns(2)
                
                with col1:
                    # Гистограмма сумм
                    fig1, ax1 = plt.subplots(figsize=(10, 6))
                    df['Сумма'].hist(bins=20, ax=ax1, alpha=0.7, color='purple')
                    ax1.set_title('Распределение сумм тендеров')
                    ax1.set_xlabel('Сумма')
                    ax1.set_ylabel('Количество')
                    st.pyplot(fig1)
                
                with col2:
                    # Суммы по разделам
                    fig2, ax2 = plt.subplots(figsize=(10, 6))
                    section_sums = df.groupby('Тип раздела')['Сумма'].sum().sort_values(ascending=False)
                    section_sums.plot(kind='bar', ax=ax2, color='coral')
                    ax2.set_title('Общая сумма по разделам')
                    ax2.set_ylabel('Сумма')
                    ax2.tick_params(axis='x', rotation=45)
                    plt.tight_layout()
                    st.pyplot(fig2)
            
            # Детальный анализ по площадкам
            st.subheader("🏢 Анализ по торговым площадкам")
            
            if 'Площадка' in df.columns:
                platform_analysis = df['Площадка'].value_counts()
                
                col1, col2 = st.columns(2)
                
                with col1:
                    st.write("**Топ площадок:**")
                    st.dataframe(platform_analysis.head(10))
                
                with col2:
                    if len(platform_analysis) > 0:
                        fig, ax = plt.subplots(figsize=(10, 6))
                        platform_analysis.head(8).plot(kind='pie', ax=ax, autopct='%1.1f%%')
                        ax.set_title('Распределение по топ площадкам')
                        ax.set_ylabel('')
                        st.pyplot(fig)
            
            # Экспорт результатов
            st.subheader("💾 Экспорт данных")
            
            col1, col2 = st.columns(2)
            
            with col1:
                if st.button("Скачать обработанные данные (CSV)"):
                    csv = df.to_csv(index=False, encoding='utf-8-sig')
                    st.download_button(
                        label="📥 Нажмите для скачивания CSV",
                        data=csv,
                        file_name="обработанные_тендеры.csv",
                        mime="text/csv"
                    )
            
            with col2:
                if st.button("Сгенерировать отчет (Excel)"):
                    excel_buffer = pd.ExcelWriter("отчет_тендеры.xlsx", engine='openpyxl')
                    df.to_excel(excel_buffer, sheet_name="Все данные", index=False)
                    
                    # Сводная таблица по разделам
                    summary = df.groupby('Тип раздела').agg({
                        '№': 'count',
                        'Сумма': ['sum', 'mean'] if 'Сумма' in df.columns else 'count'
                    }).round(2)
                    summary.to_excel(excel_buffer, sheet_name="Сводка")
                    
                    excel_buffer.close()
                    
                    with open("отчет_тендеры.xlsx", "rb") as f:
                        st.download_button(
                            label="📥 Скачать Excel отчет",
                            data=f,
                            file_name="анализ_тендеров.xlsx",
                            mime="application/vnd.ms-excel"
                        )
        
        else:
            st.error("Не удалось извлечь данные из файла. Проверьте структуру данных.")
    
    except Exception as e:
        st.error(f"Ошибка при обработке файла: {str(e)}")
        st.info("""
        **Рекомендации по формату файла:**
        - Файл должен содержать разделы с заголовками ('Мониторинг цен', 'Не участвовали' и т.д.)
        - Данные должны быть структурированы в табличном формате
        - Первые строки могут содержать метаданные
        """)

else:
    st.info("👆 Загрузите Excel файл с тендерными данными для начала анализа")
    
    # Пример ожидаемой структуры
    st.subheader("Ожидаемая структура файла:")
    st.write("""
    Файл должен содержать разделы:
    - **Мониторинг цен** - тендеры на мониторинг
    - **Не участвовали** - тендеры, в которых не участвовали  
    - **Подались** - поданные заявки
    - **Проиграли** - проигранные тендеры
    - **Отменен** - отмененные тендеры
    """)
