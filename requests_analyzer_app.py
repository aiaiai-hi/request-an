import streamlit as st
import pandas as pd
import numpy as np
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
        type=['csv', 'xlsx'],
        help="Поддерживаются файлы в форматах CSV, XLSX"
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
                df = pd.read_csv(uploaded_file, encoding='utf-8')
            elif file_extension == 'xlsx':
                df = pd.read_excel(uploaded_file)
            else:
                st.error("❌ Неподдерживаемый формат файла!")
                return
            
            # Удаляем полностью пустые строки
            df = df.dropna(how='all')
            # Удаляем строки где business_id пустой
            if 'business_id' in df.columns:
                df = df.dropna(subset=['business_id'])
            else:
                st.error("❌ В файле отсутствует столбец 'business_id'")
                return
            
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
    
    # Конвертируем даты - добавлена обработка ошибок
    date_columns = ['created_at', 'ts_from', 'ts_to']
    for col in date_columns:
        if col in df.columns:
            try:
                df[col] = pd.to_datetime(df[col], format='%d.%m.%Y', errors='coerce')
            except:
                try:
                    df[col] = pd.to_datetime(df[col], format='%Y-%m-%d', errors='coerce')
                except:
                    df[col] = pd.to_datetime(df[col], errors='coerce')
    
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
            'analyst': unique_row.get('analyst', ''),
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
