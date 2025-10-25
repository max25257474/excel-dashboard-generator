import streamlit as st
import pandas as pd
import io
import zipfile
import os
from PIL import Image
from openpyxl import load_workbook
from openpyxl.drawing.image import Image as XLImage

st.set_page_config(page_title="Генератор клієнтських файлів", layout="wide")
st.title("📁 Генератор індивідуальних Excel-файлів клієнтів")


st.sidebar.header("📂 Завантаження даних")

uploaded_file = st.sidebar.file_uploader("Завантаж свій Excel-файл", type=["xlsx", "xls"])
logo_file = st.sidebar.file_uploader("Завантаж логотип (PNG або JPG)", type=["png", "jpg", "jpeg"])

default_excel = "clients_large.xlsx"
default_logo = "logo.png"

if uploaded_file is not None:
    df = pd.read_excel(uploaded_file)
    st.success("✅ Завантажено власний Excel-файл.")
elif os.path.exists(default_excel):
    df = pd.read_excel(default_excel)
    st.success("✅ Автоматично завантажено clients_large.xlsx із локальної папки.")
else:
    st.error("❌ Не знайдено clients_large.xlsx. Завантаж файл вручну.")
    st.stop()

if logo_file is not None:
    logo_path = "user_logo.png"
    with open(logo_path, "wb") as f:
        f.write(logo_file.read())
    st.sidebar.image(logo_path, caption="Ваш логотип", use_container_width=True)
elif os.path.exists(default_logo):
    logo_path = default_logo
    st.sidebar.image(logo_path, caption="Логотип за замовчуванням", use_container_width=True)
else:
    logo_path = None
    st.sidebar.warning("⚠️ Логотип не знайдено — файли будуть без логотипу.")

st.subheader("🔍 Попередній перегляд даних")
st.dataframe(df.head(10), use_container_width=True)

columns = df.columns.tolist()
client_column = st.selectbox("Оберіть колонку з клієнтами:", columns)

if st.button("🚀 Створити індивідуальні файли"):
    clients = df[client_column].dropna().unique()
    st.info(f"Створюємо файли для {len(clients)} клієнтів...")

    zip_buffer = io.BytesIO()

    with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zipf:
        for client in clients:
            client_data = df[df[client_column] == client]

            temp_output = io.BytesIO()
            client_data.to_excel(temp_output, index=False)
            temp_output.seek(0)

            wb = load_workbook(temp_output)
            ws = wb.active

            if logo_path is not None and os.path.exists(logo_path):
                img = XLImage(logo_path)
                ws.add_image(img, "A1")

            client_output = io.BytesIO()
            wb.save(client_output)
            client_output.seek(0)


            filename = f"{client}.xlsx".replace("/", "_")
            zipf.writestr(filename, client_output.read())

    zip_buffer.seek(0)
    st.success("✅ Індивідуальні файли створено!")

    st.download_button(
        label="⬇️ Завантажити всі файли (.zip)",
        data=zip_buffer,
        file_name="clients_files.zip",
        mime="application/zip"
    )
