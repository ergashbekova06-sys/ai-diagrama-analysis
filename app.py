import streamlit as st
import pandas as pd
import matplotlib.pyplot as plt
from io import StringIO, BytesIO
from docx import Document
from docx.shared import Inches

st.title("Бақылау жұмыстарының нәтижелерін талдау")
st.write("Excel-ден алынған CSV мәтін түріндегі кестені енгізіңіз:")

csv_text = st.text_area("Кестені осында қойыңыз", height=200)

if csv_text.strip():
    try:
        # CSV жүктеу
        df = pd.read_csv(StringIO(csv_text))

        # Процент жазуларын санға айналдыру
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

        # Қажетті колонкаларды анықтау
        quality_col = None
