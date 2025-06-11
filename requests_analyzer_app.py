import streamlit as st
import pandas as pd
import numpy as np
from datetime import datetime
import io

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

# Чтение файла сразу после загрузки
if uploaded_file is not None:
    try:
        file_extension = uploaded_file.name.split('.')[-1].lower()
        if file_extension == 'csv':
            df = pd.read_csv(uploaded_file, encoding='utf-8')
        elif file_extension == 'xlsx':
            df = pd.read_excel(uploaded_file)
        else:
            st.error("❌ Неподдерживаемый формат файла!")
            df = None

        if df is not None:
            df = df.dropna(how='all')
            if 'business_id' in df.columns:
                df = df.dropna(subset=['business_id'])
                st.session_state.original_data = df
                st.success(f"✅ Файл успешно загружен! Найдено {len(df)} записей.")
            else:
                st.session_state.original_data = None
                st.error("❌ В файле отсутствует столбец 'business_id'")
    except Exception as e:
        st.session_state.original_data = None
        st.error(f"❌ Ошибка при загрузке файла: {str(e)}")
else:
    st.session_state.original_data = None

# Кнопка "Проанализировать" активна только если файл загружен
analyze_button = st.button(
    "🔍 Проанализировать",
    use_container_width=True,
    disabled=st.session_state.original_data is None
)

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
    df_sorted = df.sort_values('created_at', ascending=False)
    unique_requests = df_sorted.drop_duplicates(subset='business_id', keep='first')
    latest_records = df.groupby('business_id').apply(
        lambda x: x.loc[x['ts_from'].idxmax()] if x['ts_from'].notna().any() else x.iloc[-1]
    ).reset_index(drop=True)
    result_data = []
    for _, unique_row in unique_requests.iterrows():
        business_id = unique_row['business_id']
        latest_row = latest_records[latest_records['business_id'] == business_id].iloc[0]
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
    st.subheader("🔎 Поиск")
    search_query = st.text_input(
        "Поиск по номеру отчета (report_code) или business_id:",
        placeholder="Введите номер отчета или business_id..."
    )
    filtered_df = df.copy()
    if search_query:
        search_mask = (
            filtered_df['report_code'].astype(str).str.contains(search_query, case=False, na=False) |
            filtered_df['business_id'].astype(str).str.contains(search_query, case=False, na=False)
        )
        filtered_df = filtered_df[search_mask]
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
        if st.button("🔄 Сбросить фильтры"):
            st.rerun()
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
    st.subheader("📋 Таблица данных")
    if len(filtered_df) > 0:
        st.dataframe(
            filtered_df,
            use_container_width=True,
            hide_index=True
        )
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
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df.to_excel(writer, sheet_name='Анализ запросов', index=False)
        workbook = writer.book
        worksheet = writer.sheets['Анализ запросов']
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
