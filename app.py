import streamlit as st
import pandas as pd
import matplotlib.pyplot as plt
from io import StringIO, BytesIO
from docx import Document
from docx.shared import Inches

st.title("Анализ контрольных по видам оценивания")
st.write("Вставьте таблицу (CSV из Excel):")

csv_text = st.text_area("Вставьте таблицу сюда", height=200)

if csv_text.strip():
    try:
        df = pd.read_csv(StringIO(csv_text))

        # Преобразуем проценты
        for col in df.columns:
            if df[col].astype(str).str.contains("%").any():
                df[col] = df[col].astype(str).str.replace("%", "").str.strip().astype(float)

        st.success("Таблица обработана!")
        st.dataframe(df)

        assess_types = ["СОР 1", "СОР 2", "СОЧ"]

        # Документ Word
        document = Document()
        document.add_heading("Диаграммы по видам оценивания", level=1)

        image_buffers = []  # сюда сохраняем буферы изображений

        for assess in assess_types:
            subset = df[df["Оценивание"] == assess]

            if subset.empty:
                continue

            st.subheader(f"{assess}: Диаграммы")

            labels = subset["Класс"]
            q = subset["% Качества знаний (В + С)"]
            u = subset["% Успеваемости (Н=0)"]

            # -------- Диаграмма: два столбика --------
            fig, ax = plt.subplots(figsize=(8,4))

            x = range(len(labels))
            ax.bar([p - 0.2 for p in x], q, width=0.4, label="Качество знаний")
            ax.bar([p + 0.2 for p in x], u, width=0.4, label="Успеваемость")

            ax.set_xticks(x)
            ax.set_xticklabels(labels)
            ax.set_title(f"{assess}: Качество и Успеваемость")
            ax.set_ylabel("%")
            ax.legend()

            st.pyplot(fig)

            # сохраняем диаграмму во временный буфер для Word
            img_buf = BytesIO()
            fig.savefig(img_buf, format="png", dpi=200)
            img_buf.seek(0)
            image_buffers.append((assess, img_buf))

        # -------- Собираем Word --------
        for title, img_buf in image_buffers:
            document.add_heading(title, level=2)
            document.add_picture(img_buf, width=Inches(6))

        # создаём файл Word
        output = BytesIO()
        document.save(output)
        output.seek(0)

        st.download_button(
            label="📥 Скачать все диаграммы в Word",
            data=output,
            file_name="Диаграммы_анализ.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        )

    except Exception as e:
        st.error(f"Ошибка: {e}")
