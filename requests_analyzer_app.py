import streamlit as st
import pandas as pd
import numpy as np
import openpyxl
from datetime import datetime, timedelta
import io

def main():
    st.set_page_config(
        page_title="Анализатор запросов",
        page_icon="📊",
        layout="wide"
    )
    
    st.title("📊 Анализатор запросов и стадий рассмотрения")
    st.markdown("---")
    
    # Инициализация session state
    if 'processed_data' not in st.session_state:
        st.session_state.processed_data = None
    if 'original_data' not in st.session_state:
        st.session_state.original_data = None
    
    # Форма загрузки файла
    st.subheader("📁 Загрузка файла для анализа")
    
    uploaded_file = st.file_uploader(
        "Выберите файл с данными о запросах",
        type=['csv', 'xls', 'xlsx'],
        help="Поддерживаются файлы в форматах CSV, XLS, XLSX"
    )
    
    col1, col2 = st.columns([1, 1])
    
    with col1:
        load_button = st.button("📤 Загрузить файл", use_container_width=True)
    
    with col2:
        analyze_button = st.button("🔍 Проанализировать", use_container_width=True, disabled=uploaded_file is None)
    
    # Обработка загрузки файла
    if load_button and uploaded_file:
        try:
            # Определяем тип файла и читаем соответствующим образом
            file_extension = uploaded_file.name.split('.')[-1].lower()
            
            if file_extension == 'csv':
                # Чтение CSV файла
                df = pd.read_csv(uploaded_file, encoding='utf-8')
            elif file_extension in ['xls', 'xlsx']:
                # Чтение Excel файла
                df = pd.read_excel(uploaded_file, engine='openpyxl' if file_extension == 'xlsx' else None)
            else:
                st.error("❌ Неподдерживаемый формат файла!")
                return
            
            # Удаляем полностью пустые строки
            df = df.dropna(how='all')
            # Удаляем строки где business_id пустой
            df = df.dropna(subset=['business_id'])
            
            st.session_state.original_data = df
            st.success(f"✅ Файл успешно загружен! Найдено {len(df)} записей.")
            
        except Exception as e:
            st.error(f"❌ Ошибка при загрузке файла: {str(e)}")
    
    # Обработка анализа данных
    if analyze_button and st.session_state.original_data is not None:
        try:
            processed_data = process_data(st.session_state.original_data)
            st.session_state.processed_data = processed_data
            st.success("✅ Данные успешно обработаны!")
            
        except Exception as e:
            st.error(f"❌ Ошибка при обработке данных: {str(e)}")
    
    # Отображение результатов
    if st.session_state.processed_data is not None:
        display_results(st.session_state.processed_data)

def process_data(df):
    """Обработка данных согласно требованиям"""
    
    # Конвертируем даты
    date_columns = ['created_at', 'ts_from', 'ts_to']
    for col in date_columns:
        if col in df.columns:
            df[col] = pd.to_datetime(df[col], format='%d.%m.%Y', errors='coerce')
    
    # Сортируем по created_at от новых к старым
    df_sorted = df.sort_values('created_at', ascending=False)
    
    # Получаем уникальные запросы по business_id (берем первый после сортировки - самый новый)
    unique_requests = df_sorted.drop_duplicates(subset='business_id', keep='first')
    
    # Для каждого business_id находим последнюю строку для расчета дней в работе
    latest_records = df.groupby('business_id').apply(
        lambda x: x.loc[x['ts_from'].idxmax()] if x['ts_from'].notna().any() else x.iloc[-1]
    ).reset_index(drop=True)
    
    # Создаем итоговую таблицу
    result_data = []
    
    for _, unique_row in unique_requests.iterrows():
        business_id = unique_row['business_id']
        
        # Находим соответствующую последнюю запись для расчета дней
        latest_row = latest_records[latest_records['business_id'] == business_id].iloc[0]
        
        # Рассчитываем дни в работе
        if pd.notna(latest_row['ts_from']):
            days_in_work = (datetime.now() - latest_row['ts_from']).days
        else:
            days_in_work = 0
        
        result_data.append({
            'business_id': int(business_id),
            'created_at': unique_row['created_at'].strftime('%d.%m.%Y') if pd.notna(unique_row['created_at']) else '',
            'дней_в_работе': days_in_work,
            'form_type_report': unique_row.get('form_type_report', ''),
            'report_code': unique_row.get('report_code', ''),
            'report_name': unique_row.get('report_name', ''),
            'current_stage': unique_row.get('current_stage', ''),
            'ts_from': latest_row['ts_from'].strftime('%d.%m.%Y') if pd.notna(latest_row['ts_from']) else '',
            'analyst': unique_row.get('Analyst', ''),
            'request_owner': unique_row.get('request_owner', ''),
            'request_owner_ssp': unique_row.get('request_owner_ssp', '')
        })
    
    return pd.DataFrame(result_data)

def display_results(df):
    """Отображение результатов с фильтрами и поиском"""
    
    st.markdown("---")
    st.subheader("🔍 Результаты анализа")
    
    # Строка поиска
    st.subheader("🔎 Поиск")
    search_query = st.text_input(
        "Поиск по номеру отчета (report_code) или business_id:",
        placeholder="Введите номер отчета или business_id..."
    )
    
    # Применяем поиск
    filtered_df = df.copy()
    if search_query:
        search_mask = (
            filtered_df['report_code'].astype(str).str.contains(search_query, case=False, na=False) |
            filtered_df['business_id'].astype(str).str.contains(search_query, case=False, na=False)
        )
        filtered_df = filtered_df[search_mask]
    
    # Фильтры
    st.subheader("🔧 Фильтры")
    
    col1, col2, col3, col4 = st.columns(4)
    
    with col1:
        form_types = ['Все'] + sorted(df['form_type_report'].dropna().unique().tolist())
        selected_form_type = st.selectbox("Тип формы отчета:", form_types)
        
        analysts = ['Все'] + sorted(df['analyst'].dropna().unique().tolist())
        selected_analyst = st.selectbox("Аналитик:", analysts)
    
    with col2:
        stages = ['Все'] + sorted(df['current_stage'].dropna().unique().tolist())
        selected_stage = st.selectbox("Текущая стадия:", stages)
        
        owners = ['Все'] + sorted(df['request_owner'].dropna().unique().tolist())
        selected_owner = st.selectbox("Владелец запроса:", owners)
    
    with col3:
        owner_ssps = ['Все'] + sorted(df['request_owner_ssp'].dropna().unique().tolist())
        selected_owner_ssp = st.selectbox("Владелец ССП:", owner_ssps)
        
        min_days = st.number_input("Мин. дней в работе:", min_value=0, value=0)
    
    with col4:
        max_days = st.number_input("Макс. дней в работе:", min_value=0, value=1000)
        
        # Кнопка сброса фильтров
        if st.button("🔄 Сбросить фильтры"):
            st.rerun()
    
    # Применяем фильтры
    if selected_form_type != 'Все':
        filtered_df = filtered_df[filtered_df['form_type_report'] == selected_form_type]
    
    if selected_stage != 'Все':
        filtered_df = filtered_df[filtered_df['current_stage'] == selected_stage]
    
    if selected_analyst != 'Все':
        filtered_df = filtered_df[filtered_df['analyst'] == selected_analyst]
    
    if selected_owner != 'Все':
        filtered_df = filtered_df[filtered_df['request_owner'] == selected_owner]
    
    if selected_owner_ssp != 'Все':
        filtered_df = filtered_df[filtered_df['request_owner_ssp'] == selected_owner_ssp]
    
    filtered_df = filtered_df[
        (filtered_df['дней_в_работе'] >= min_days) & 
        (filtered_df['дней_в_работе'] <= max_days)
    ]
    
    # Статистика
    col1, col2, col3, col4 = st.columns(4)
    with col1:
        st.metric("📊 Всего записей", len(df))
    with col2:
        st.metric("🔍 После фильтрации", len(filtered_df))
    with col3:
        if len(filtered_df) > 0:
            avg_days = filtered_df['дней_в_работе'].mean()
            st.metric("📅 Среднее дней в работе", f"{avg_days:.1f}")
        else:
            st.metric("📅 Среднее дней в работе", "0")
    with col4:
        if len(filtered_df) > 0:
            max_days_value = filtered_df['дней_в_работе'].max()
            st.metric("⏰ Максимум дней", max_days_value)
        else:
            st.metric("⏰ Максимум дней", "0")
    
    # Отображение таблицы
    st.subheader("📋 Таблица данных")
    
    if len(filtered_df) > 0:
        # Настройка отображения столбцов
        column_config = {
            'business_id': st.column_config.NumberColumn('Business ID', format='%d'),
            'created_at': st.column_config.TextColumn('Дата создания'),
            'дней_в_работе': st.column_config.NumberColumn('Дней в работе', format='%d'),
            'form_type_report': st.column_config.TextColumn('Тип формы'),
            'report_code': st.column_config.TextColumn('Код отчета'),
            'report_name': st.column_config.TextColumn('Название отчета'),
            'current_stage': st.column_config.TextColumn('Текущая стадия'),
            'ts_from': st.column_config.TextColumn('Дата начала'),
            'analyst': st.column_config.TextColumn('Аналитик'),
            'request_owner': st.column_config.TextColumn('Владелец запроса'),
            'request_owner_ssp': st.column_config.TextColumn('Владелец ССП')
        }
        
        st.dataframe(
            filtered_df,
            use_container_width=True,
            column_config=column_config,
            hide_index=True
        )
        
        # Кнопка экспорта в Excel
        if st.button("📥 Скачать в Excel", use_container_width=True):
            excel_data = create_excel_download(filtered_df)
            st.download_button(
                label="💾 Загрузить Excel файл",
                data=excel_data,
                file_name=f"requests_analysis_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
    else:
        st.warning("⚠️ Нет данных для отображения. Попробуйте изменить фильтры.")

def create_excel_download(df):
    """Создание Excel файла для скачивания"""
    output = io.BytesIO()
    
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df.to_excel(writer, sheet_name='Анализ запросов', index=False)
        
        # Получаем workbook и worksheet для форматирования
        workbook = writer.book
        worksheet = writer.sheets['Анализ запросов']
        
        # Автоподбор ширины столбцов
        for column in worksheet.columns:
            max_length = 0
            column_letter = column[0].column_letter
            for cell in column:
                try:
                    if len(str(cell.value)) > max_length:
                        max_length = len(str(cell.value))
                except:
                    pass
            adjusted_width = min(max_length + 2, 50)
            worksheet.column_dimensions[column_letter].width = adjusted_width
    
    output.seek(0)
    return output.getvalue()

if __name__ == "__main__":
    main()
