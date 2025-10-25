import streamlit as st
import pandas as pd
import io
import zipfile
import os
from PIL import Image
from openpyxl import load_workbook
from openpyxl.drawing.image import Image as XLImage

st.set_page_config(page_title="Генератор файлів клієнтів", layout="wide")
st.title("📁 Генератор індивідуальних Excel-файлів клієнтів")


uploaded_file = st.file_uploader("Завантаж загальний Excel-файл", type=["xlsx", "xls"])

logo_file = st.file_uploader("Завантаж логотип (PNG або JPG)", type=["png", "jpg", "jpeg"])

if uploaded_file is not None:
    df = pd.read_excel(uploaded_file)
    st.success(f"✅ Файл завантажено ({len(df)} записів)")
    st.dataframe(df.head(), use_container_width=True)


    columns = df.columns.tolist()
    client_column = st.selectbox("Оберіть колонку з іменами клієнтів:", columns)

 
    if st.button("🚀 Створити індивідуальні файли"):
        clients = df[client_column].dropna().unique()
        st.info(f"Створюємо файли для {len(clients)} клієнтів...")

        zip_buffer = io.BytesIO()

        with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zipf:
            for client in clients:
                client_data = df[df[client_column] == client]


                output = io.BytesIO()
                client_data.to_excel(output, index=False)
                output.seek(0)


                wb = load_workbook(output)
                ws = wb.active


                if logo_file is not None:
                    logo_image = Image.open(logo_file)
                    temp_path = f"temp_logo.png"
                    logo_image.save(temp_path)
                    img = XLImage(temp_path)
                    ws.add_image(img, "A1")
                    os.remove(temp_path)


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
else:
    st.info("👆 Завантаж таблицю клієнтів для початку роботи.")
