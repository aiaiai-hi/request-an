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
        st.error(f
