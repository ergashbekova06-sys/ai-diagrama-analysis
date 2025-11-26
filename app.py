import streamlit as st
import pandas as pd
import matplotlib.pyplot as plt
from io import StringIO, BytesIO
from docx import Document
from docx.shared import Inches

st.title("БЖБ ЖӘНЕ ТЖБ ТАЛДАУ ДИАГРАММАЛАРЫ")
st.write(" 19 ЖАЛПЫ БІЛІМ БЕРЕТІН МЕКТЕП КММ")

csv_text = st.text_area("Excel CSV электрондық кестесін осы жерге жазыңыз", height=200)

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

        st.success("Кесте жүктелді!")
        st.dataframe(df)

        # Ищем колонки автоматически
        quality_col = None
        success_col = None

        for col in df.columns:
            col_low = col.lower()

            if "біл" in col_low:
                quality_col = col
            if "үлгер" in col_low:
                success_col = col

        if not quality_col or not success_col:
            st.error("Не найдены колонки 'качество' или 'успеваемость'.")
            st.stop()

        st.info(f"Колонка білім сапасы: **{quality_col}**")
        st.info(f"Колонка үлгерімі: **{success_col}**")

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

            ax.bar([p - 0.2 for p in x], q, width=0.4, label="Біліім сапасы")
            ax.bar([p + 0.2 for p in x], u, width=0.4, label="Үлгерімі")

            ax.set_xticks(x)
            ax.set_xticklabels(labels)
            ax.set_title(f"{assess}: Сапа және үлгерім")
            ax.set_ylabel("%")
            ax.legend()

            st.pyplot(fig)

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
            "📥 Word бағдарламасына диаграммаларды жүктеп алыңыз",
            data=output,
            file_name="БЖБ ЖӘНЕ ТЖБ ТАЛДАУ ДИАГРАММАЛАРЫ.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        )

    except Exception as e:
        st.error(f"Ошибка: {e}")
