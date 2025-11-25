import streamlit as st
import pandas as pd
import matplotlib.pyplot as plt
from io import StringIO

st.title("Анализ контрольных работ — диаграммы")

st.write("Вставьте таблицу в формате CSV (как вы копируете из Excel):")

# Ввод CSV текста
csv_text = st.text_area("Вставьте таблицу сюда", height=200)

if csv_text:
    try:
        # Преобразуем текст в DataFrame
        df = pd.read_csv(StringIO(csv_text))

        # Приводим процентные колонки к числам
        for col in df.columns:
            if df[col].dtype == object and df[col].astype(str).str.contains("%").any():
                df[col] = df[col].astype(str).str.replace("%", "").str.strip().astype(float)

        st.success("Таблица успешно загружена!")
        st.write(df)

        # Выбор показателя для графика
        numeric_columns = df.select_dtypes(include=['float64', 'int64']).columns
        metric = st.selectbox("Выберите показатель для диаграммы:", numeric_columns)

        # Построение диаграммы
        fig, ax = plt.subplots(figsize=(8, 4))
        ax.bar(df["Класс"] + " " + df["Оценивание"], df[metric])
        ax.set_xticklabels(df["Класс"] + " " + df["Оценивание"], rotation=45, ha="right")
        ax.set_ylabel(metric)
        ax.set_title(f"Диаграмма: {metric}")

        st.pyplot(fig)

    except Exception as e:
        st.error(f"Ошибка при обработке: {e}")
