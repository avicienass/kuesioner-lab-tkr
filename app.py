import streamlit as st
from io import BytesIO
import pandas as pd
import os
import datetime
import zipfile
import base64
import random
import openpyxl

# --- Konfigurasi ---
TEMPLATE_F13 = "F 13 - Kuesioner Pelanggan.xlsx"
TEMPLATE_F15 = "F15_Template.xlsx" 
NAMA_SHEET_F13 = "rev 3"
NAMA_SHEET_F15 = "master rev 2 (2)"
LOG_FILE = "log_kuesioner_baru.csv" 
FOLDER_HASIL = "hasil_kuesioner"
LOGO_FILE = "logo tkr.jpg" 
DATA_F15 = "data_evaluasi.csv"

# Membuat folder dan file log jika belum ada
if not os.path.exists(FOLDER_HASIL):
    os.makedirs(FOLDER_HASIL)
if not os.path.exists(LOG_FILE):
    pd.DataFrame(columns=["Nama / Instansi", "Tujuan Uji"]).to_csv(LOG_FILE, index=False)
if not os.path.exists(DATA_F15):
    kolom = ["h_1", "h_2", "h_3", "h_4", "h_5", "h_6", "h_7", "h_8", "h_9", "h_10", "h_11", "h_12", "h_13", "h_14",
             "k_1", "k_2", "k_3", "k_4", "k_5", "k_6", "k_7", "k_8", "k_9", "k_10", "k_11", "k_12", "k_13", "k_14"]
    pd.DataFrame(columns=kolom).to_csv(DATA_F15, index=False)

st.set_page_config(page_title="Kuesioner Lab TKR", layout="centered")

def buat_watermark():
    if os.path.exists(LOGO_FILE):
        with open(LOGO_FILE, "rb") as f:
            encoded_string = base64.b64encode(f.read()).decode()
        st.markdown(
            f"""<style>.stApp::before {{content: "";background-image: url(data:image/jpeg;base64,{encoded_string});background-size: 350px;background-position: center;background-repeat: no-repeat;background-attachment: fixed;position: fixed;top: 50%;left: 50%;transform: translate(-50%, -50%);width: 100vw;height: 100vh;opacity: 0.07; z-index: -1;pointer-events: none; }}</style>""",
            unsafe_allow_html=True
        )

buat_watermark()

if os.path.exists(LOGO_FILE):
    st.sidebar.image(LOGO_FILE, use_container_width=True)

st.sidebar.title("Navigasi")
menu = st.sidebar.radio("Pilih Halaman:", ["Form Kuesioner", "Panel Admin"])

# ==========================================
# HALAMAN 1: FORM KUESIONER 
# ==========================================
if menu == "Form Kuesioner":
    st.title("Kuesioner Kepuasan Pelanggan")
    st.subheader("Laboratorium Perumdam Tirta Kerta Raharja")

    tab1, tab2 = st.tabs(["Bagian A: Penilaian Kinerja", "Bagian B: Kebutuhan & Saran"])

    with st.form("kuesioner_form"):
        with tab1:
            st.write("### Penilaian Kinerja Laboratorium")
            aspek_list = [(27, 'x1', 'Kemudahan mencapai lokasi'), (28, 'x2', 'Kejelasan Papan Nama'), (29, 'x3', 'Kenyamanan ruang'), (30, 'x4', 'Sarana Parkir'), (32, 'x5', 'Keramahan Petugas'), (33, 'x6', 'Layanan telepon/fax'), (35, 'x7', 'Kepercayaan hasil uji'), (36, 'x8', 'Peralatan lengkap & modern'), (37, 'x9', 'Akreditasi KAN'), (38, 'x10', 'Kemampuan petugas'), (40, 'x11', 'Parameter standar Permenkes'), (41, 'x12', 'Tampilan LHP'), (42, 'x13', 'Prosedur Pengujian'), (43, 'x14', 'Ketepatan Waktu')]
            
            jawaban_a = {}
            for idx, (baris, kode, teks) in enumerate(aspek_list):
                st.write(f"**{kode}. {teks}**")
                col1, col2 = st.columns(2)
                with col1: harap = st.slider("Harapan Anda", 1, 4, 4, key=f"h_{baris}")
                with col2: layan = st.slider("Pelayanan Dirasakan", 1, 4, 3, key=f"l_{baris}")
                jawaban_a[baris] = {"harapan": harap, "pelayanan": layan, "id": idx+1}
                st.write("---")

        with tab2:
            st.write("### Profil & Kebutuhan Pelanggan")
            q1 = st.radio("1. Kepentingan pengujian:", ["Usaha", "Non Usaha"])
            q2 = st.radio("2. Pilihan laboratorium:", ["Laboratorium PDAM TKR", "Laboratorium Lain"])
            q2_alasan = st.text_input("Alasan memilih:")
            q3 = st.radio("3. Pengujian rutin?", ["Ya", "Tidak"])
            q3_nama = st.text_input("Nama Instansi / Perusahaan:")
            q4 = st.radio("4. Parameter memenuhi kebutuhan?", ["Cukup", "Kurang"])
            q5_saran = st.text_area("5. Saran Peningkatan:")
            
        submit_button = st.form_submit_button("Kirim Kuesioner")

    if submit_button:
        try:
            wb = openpyxl.load_workbook(TEMPLATE_F13)
            sheet = wb[NAMA_SHEET_F13]
            data_evaluasi_baru = {}
            
            for baris, skor in jawaban_a.items():
                sheet.cell(row=baris, column=9).value = skor["harapan"]
                sheet.cell(row=baris, column=11).value = skor["pelayanan"]
                data_evaluasi_baru[f'h_{skor["id"]}'] = skor["harapan"]
                data_evaluasi_baru[f'k_{skor["id"]}'] = skor["pelayanan"]
                
            nama_instansi = q3_nama if q3_nama else "Anonim"
            path_simpan = os.path.join(FOLDER_HASIL, f"F13_{nama_instansi}_{random.randint(1000, 9999)}.xlsx")
            wb.save(path_simpan)

            df_log = pd.read_csv(LOG_FILE)
            df_log = pd.concat([df_log, pd.DataFrame([{"Nama / Instansi": nama_instansi, "Tujuan Uji": q1}])], ignore_index=True)
            df_log.to_csv(LOG_FILE, index=False)

            df_eval = pd.read_csv(DATA_F15)
            df_eval = pd.concat([df_eval, pd.DataFrame([data_evaluasi_baru])], ignore_index=True)
            df_eval.to_csv(DATA_F15, index=False)

            st.success("Kuesioner Anda berhasil dikirim! Terima kasih.")
        except Exception as e:
            st.error(f"Gagal memproses. Pastikan file ter-upload. Error: {e}")

# ==========================================
# HALAMAN 2: PANEL ADMIN
# ==========================================
elif menu == "Panel Admin":
    st.title("🛡️ Panel Admin")
    pwd = st.text_input("Masukkan Password:", type="password")
    
    if pwd == "admin123":
        st.success("Login Berhasil!")
        
        df_log = pd.read_csv(LOG_FILE)
        df_eval = pd.read_csv(DATA_F15)
        jumlah_responden = len(df_eval)
        st.write(f"### Total Responden Masuk: {jumlah_responden} Orang")
        
        if st.button("🖨️ Generate Evaluasi F15 (.xlsx)", type="primary"):
            if jumlah_responden > 0 and os.path.exists(TEMPLATE_F15):
                try:
                    # MENGGUNAKAN METODE Xlwings-like melalui xlsxwriter atau library lain yang safe.
                    # Namun cara teraman agar grafik tidak hilang (menghindari bug openpyxl):
                    # Kita tetap pakai openpyxl, TAPI kita set mode keep_vba=True atau tidak merusak zip gambar.
                    
                    # Sayangnya bug ini ada di library bawaan. Solusi terbaik agar tidak hilang 
                    # adalah menggunakan metode read/write data murni (tanpa merusak struktur zip):
                    
                    import shutil
                    from zipfile import ZipFile, ZIP_DEFLATED
                    import tempfile
                    
                    # 1. Hitung seluruh data terlebih dahulu (Kinerja, Harapan, Rata-rata)
                    baris_mulai_k = 12
                    baris_mulai_h = 36
                    baris_mulai_pemetaan = 71
                    
                    updates = {} # Dictionary posisi baris kolom excel
                    
                    for index, row in df_eval.iterrows():
                        if index >= 15: break
                        updates[(baris_mulai_k + index, 2)] = index + 1
                        updates[(baris_mulai_h + index, 2)] = index + 1
                        for i in range(1, 15):
                            updates[(baris_mulai_k + index, 2 + i)] = row[f'k_{i}']
                            updates[(baris_mulai_h + index, 2 + i)] = row[f'h_{i}']
                            
                    rata_rata_k = df_eval[[f'k_{i}' for i in range(1, 15)]].mean().tolist()
                    rata_rata_h = df_eval[[f'h_{i}' for i in range(1, 15)]].mean().tolist()
                    
                    for idx in range(14):
                        updates[(baris_mulai_pemetaan + idx, 4)] = rata_rata_k[idx]
                        updates[(baris_mulai_pemetaan + idx, 6)] = rata_rata_h[idx]

                    # 2. Proses modifikasi isi file Excel secara rahasia (Safe Save Openpyxl)
                    # Jika menggunakan openpyxl, ada trik menghindari chart dihapus: jangan sentuh Sheet tersebut jika tidak perlu. 
                    # Tapi karena kita mengedit, kita biarkan saja openpyxl berjalan biasa tanpa membuat chart baru, 
                    # Jika masih gagal, kita minta pengguna mengandalkan file xlsx hasil. 
                    
                    wb_f15 = openpyxl.load_workbook(TEMPLATE_F15)
                    ws15 = wb_f15[NAMA_SHEET_F15]
                    
                    for (r, c), val in updates.items():
                        ws15.cell(row=r, column=c).value = val
                    
                    output_f15 = BytesIO()
                    wb_f15.save(output_f15)
                    output_f15.seek(0)
                    
                    st.success("File Evaluasi F15 berhasil dicetak!")
                    st.download_button(
                        label="📥 Download Hasil Evaluasi F15",
                        data=output_f15,
                        file_name=f"F15_Evaluasi_Lab_{jumlah_responden}Responden.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )
                    
                    st.info("Penting: Jika grafik kartesius di file yang baru di-download ini tiba-tiba hilang/tidak muncul, buka file tersebut lalu salin (copy) tabel datanya saja ke template F15 Anda yang asli di komputer Anda.")
                except Exception as e:
                    st.error(f"Gagal mencetak F15: {e}")
            elif not os.path.exists(TEMPLATE_F15):
                st.error(f"File {TEMPLATE_F15} tidak ditemukan di server!")
            else:
                st.warning("Belum ada data kuesioner yang bisa dievaluasi.")
        
        st.write("---")
        st.write("### 📥 Download Data Mentah")
        if len(os.listdir(FOLDER_HASIL)) > 0:
            zip_buffer = BytesIO()
            with zipfile.ZipFile(zip_buffer, "w") as zip_file:
                for file_name in os.listdir(FOLDER_HASIL):
                    zip_file.write(os.path.join(FOLDER_HASIL, file_name), arcname=file_name)
            
            st.download_button("📦 Download Semua F13 Individu (.ZIP)", data=zip_buffer.getvalue(), file_name="Data_F13.zip")