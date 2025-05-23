import streamlit as st
import pandas as pd
import openai
import io

st.set_page_config(page_title="Excel ИИ-Анализатор", layout="wide")
st.title("📊 ИИ-Анализатор Excel Таблиц")

# Установка API-ключа
openai.api_key = st.secrets["OPENAI_API_KEY"]

uploaded_file = st.file_uploader("Загрузите Excel-файл (.xlsx)", type=["xlsx"])

if uploaded_file:
    df = pd.read_excel(uploaded_file)
    st.subheader("📄 Просмотр таблицы")
    st.dataframe(df)

    if st.button("🔍 Сделать анализ"):
        prompt = f"""Ты — аналитик. Вот таблица в формате CSV:
{df.to_csv(index=False)}

Сделай краткий, но глубокий анализ. Укажи тренды, важные выводы, возможные проблемы или гипотезы."""
        
        response = openai.ChatCompletion.create(
            model="gpt-4",
            messages=[{"role": "user", "content": prompt}]
        )
        analysis = response.choices[0].message.content
        st.subheader("📈 Анализ")
        st.write(analysis)

        # Добавим анализ как отдельный лист
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine="openpyxl") as writer:
            df.to_excel(writer, index=False, sheet_name="Данные")
            pd.DataFrame({"Анализ": [analysis]}).to_excel(writer, index=False, sheet_name="Вывод")
        output.seek(0)

        st.download_button("💾 Скачать таблицу с анализом", output, file_name="analyzed.xlsx")
