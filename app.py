# app.py

import os
import zipfile
import pandas as pd
import streamlit as st

def sanitize_filename(name):
    return "".join(c for c in name if c not in '\\/:*?"<>|')

def process_files(dwf_files, excel_file):
    renamed = []
    original = []
    report = []

    try:
        df = pd.read_excel(excel_file, sheet_name='Information')
        df = df[['DWG_No', 'DWG_Title']].dropna()
        mapping = {
            str(no).strip(): str(title).strip()
            for no, title in zip(df['DWG_No'], df['DWG_Title'])
        }
    except Exception as e:
        st.error(f"Excel okunamadı: {e}")
        return None

    output = "renamed_output"
    os.makedirs(output, exist_ok=True)

    for file in dwf_files:
        base = os.path.splitext(file.name)[0].strip()
        title = mapping.get(base)
        if title:
            name = f"{base}+{sanitize_filename(title)}.dwf"
            renamed.append(name)
        else:
            name = file.name
            original.append(name)

        path = os.path.join(output, name)
        i = 1
        while os.path.exists(path):
            name_mod = f"{os.path.splitext(name)[0]}_{i}.dwf"
            path = os.path.join(output, name_mod)
            i += 1

        with open(path, "wb") as f:
            f.write(file.read())

    zipname = "renamed_dwfs.zip"
    with zipfile.ZipFile(zipname, "w") as zipf:
        for fname in os.listdir(output):
            zipf.write(os.path.join(output, fname), arcname=fname)

    report.append(f"{len(renamed)} dosya yeniden adlandırıldı.")
    report.append(f"{len(original)} dosya orijinal adıyla kaydedildi.")
    return zipname, report

st.title("📂 DWF Dosya Yeniden Adlandırıcı")

excel = st.file_uploader("📄 Excel dosyası (Information sayfası içeriyor)", type=["xls", "xlsx"])
dwfs = st.file_uploader("📁 DWF dosyalarını yükleyin", type=["dwf"], accept_multiple_files=True)

if st.button("🚀 İşlemi Başlat") and excel and dwfs:
    result = process_files(dwfs, excel)
    if result:
        zipfile_path, report = result
        st.success("\n".join(report))
        with open(zipfile_path, "rb") as f:
            st.download_button("📦 ZIP dosyasını indir", f, file_name="renamed_dwfs.zip")
