import streamlit as st
from openpyxl import Workbook
from datetime import date, timedelta
import pandas as pd
import io

st.title("Kadın Doğum Staj Rotasyon Sistemi")

start_date = st.date_input("Başlangıç Tarihi", value=date(2026, 6, 29))
days_count = st.number_input("Gün Sayısı", 28)

groups = ["Grup1","Grup2","Grup3","Grup4","Grup5"]

data = []

for i in range(days_count):
    d = start_date + timedelta(days=i)

    data.append({
        "Tarih": d,
        "Nöbet": groups[i % 5],
        "İzinli": "" if i == 0 else groups[(i-1) % 5],
        "Gökhan": "KURA",
        "Onko": "Kadın",
        "NST": "Kadın"
    })

df = pd.DataFrame(data)

st.dataframe(df)

if st.button("Excel indir"):

    wb = Workbook()
    ws = wb.active
    ws.append(list(df.columns))

    for row in df.values:
        ws.append(list(row))

    output = io.BytesIO()
    wb.save(output)
    output.seek(0)

    st.download_button(
        "İndir",
        output,
        "rotasyon.xlsx",
        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
