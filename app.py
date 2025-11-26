import streamlit as st
import pandas as pd
import matplotlib.pyplot as plt
from io import StringIO, BytesIO
from docx import Document
from docx.shared import Inches

st.title("Анализ контрольных работ")
st.write("Вставьте таблицу (CSV из Excel):")

csv_text = st.text_area("Вставьте таблицу сюда", height=200)

if csv_text.strip():
    try:
        # Загружаем CSV-текст
        df = pd.read_csv(StringIO(csv_text))

        # Преобразуем проценты в числа
        for col in df.columns:
            if df[col].astype(str).str.contains("%").any():
                df[col] = (
                    df[col]
                    .astype(str)
                    .str.replace("%", "")
                    .str.replace(",", ".")
                    .str.strip()
                    .astype(float)
                )

        st.success("Таблица загружена!")
        st.dataframe(df)

        # Ищем колонки автоматически
        quality_col = None
        success_col = None

        for col in df.columns:
            col_low = col.lower()

            if "кач" in col_low:
                quality_col = col
            if "успе" in col_low:
                success_col = col

        if not quality_col or not success_col:
            st.error("Не найдены колонки 'качество' или 'успеваемость'.")
            st.stop()

        st.info(f"Колонка качества: **{quality_col}**")
        st.info(f"Колонка успеваемости: **{success_col}**")

        # Типы оценивания
        assess_types = ["СОР 1", "СОР 2", "СОЧ"]

        # Для Word
        document = Document()
        document.add_heading("Анализ контрольных работ", level=1)
        image_buffers = []

        # ---------------- Диаграммы ----------------
        for assess in assess_types:
            subset = df[df["Оценивание"].str.contains(assess, case=False, na=False)]

            if subset.empty:
                continue

            st.subheader(f"{assess}: Диаграммы")

            labels = subset["Класс"]
            q = subset[quality_col]
            u = subset[success_col]

            fig, ax = plt.subplots(figsize=(8, 4))
            x = range(len(labels))

            ax.bar([p - 0.2 for p in x], q, width=0.4, label="Качество знаний")
            ax.bar([p + 0.2 for p in x], u, width=0.4, label="Успеваемость")
           
            
            ax.set_xticks(x)
            ax.set_xticklabels(labels)
            ax.set_title(f"{assess}: Качество и Успеваемость")
            ax.set_ylabel("%")
            ax.legend()
            # Сохраняем в память для Word
            img_buf = BytesIO()
            fig.savefig(img_buf, format="png", dpi=200)
            img_buf.seek(0)
            image_buffers.append((assess, img_buf))

        # ---------------- Word ----------------
        for title, img_buf in image_buffers:
            document.add_heading(title, level=2)
            document.add_picture(img_buf, width=Inches(6))

        output = BytesIO()
        document.save(output)
        output.seek(0)

        st.download_button(
            "📥 Скачать диаграммы в Word",
            data=output,
            file_name="Анализ_контрольных_работ.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        )

    except Exception as e:
        st.error(f"Ошибка: {e}")
