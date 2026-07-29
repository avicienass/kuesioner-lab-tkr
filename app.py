import streamlit as st
import openpyxl
from io import BytesIO
import pandas as pd
import os
import datetime
import zipfile
import base64

# --- Konfigurasi ---
TEMPLATE_FILE = "F 13 - Kuesioner Pelanggan.xlsx"
NAMA_SHEET = "rev 3"
LOG_FILE = "log_responden.csv"
FOLDER_HASIL = "hasil_kuesioner"
LOGO_FILE = "logo tkr.jpg" # Menghubungkan file logo Anda

# Membuat folder dan file log jika belum ada di server
if not os.path.exists(FOLDER_HASIL):
    os.makedirs(FOLDER_HASIL)
if not os.path.exists(LOG_FILE):
    pd.DataFrame(columns=["Waktu Isi", "Nama / Instansi", "Tujuan Uji"]).to_csv(LOG_FILE, index=False)

st.set_page_config(page_title="Kuesioner Lab TKR", layout="centered")

# --- FUNGSI MEMBUAT WATERMARK ---
def buat_watermark():
    if os.path.exists(LOGO_FILE):
        with open(LOGO_FILE, "rb") as f:
            encoded_string = base64.b64encode(f.read()).decode()
        
        st.markdown(
            f"""
            <style>
            .stApp::before {{
                content: "";
                background-image: url(data:image/jpeg;base64,{encoded_string});
                background-size: 350px;
                background-position: center;
                background-repeat: no-repeat;
                background-attachment: fixed;
                position: fixed;
                top: 50%;
                left: 50%;
                transform: translate(-50%, -50%);
                width: 100vw;
                height: 100vh;
                opacity: 0.07; /* Tingkat transparansi watermark (7%) */
                z-index: -1;
                pointer-events: none; /* Agar tidak menghalangi klik */
            }}
            </style>
            """,
            unsafe_allow_html=True
        )

# Jalankan fungsi watermark
buat_watermark()

# --- MENU SAMPING (SIDEBAR) ---
# Menampilkan logo PERUMDAM yang solid (tidak transparan) di menu samping
if os.path.exists(LOGO_FILE):
    st.sidebar.image(LOGO_FILE, use_container_width=True)

st.sidebar.title("Navigasi")
menu = st.sidebar.radio("Pilih Halaman:", ["Form Kuesioner", "Panel Admin"])

# ==========================================
# HALAMAN 1: FORM KUESIONER (UNTUK PELANGGAN)
# ==========================================
if menu == "Form Kuesioner":
    st.title("Kuesioner Kepuasan Pelanggan")
    st.subheader("Laboratorium Perumdam Tirta Kerta Raharja")
    st.write("Terima kasih atas kesediaan Anda menjawab pertanyaan berikut. Berikan nilai **1 (Tidak Baik/Tidak Penting)** hingga **4 (Sangat Baik/Sangat Penting)**.")

    tab1, tab2 = st.tabs(["Bagian A: Penilaian Kinerja", "Bagian B: Kebutuhan & Saran"])

    with st.form("kuesioner_form"):
        with tab1:
            st.write("### Penilaian Kinerja Laboratorium")
            aspek_list = [
                (27, 'x1. Kemudahan mencapai lokasi laboratorium'),
                (28, 'x2. Kejelasan Papan Nama Gedung'),
                (29, 'x3. Kenyamanan dan kebersihan ruang'),
                (30, 'x4. Sarana Tempat Parkir'),
                (32, 'x5. Keramahan Petugas'),
                (33, 'x6. Kemudahan Layanan melalui telepon/ fax'),
                (35, 'x7. Kepercayaan terhadap hasil pengujian'),
                (36, 'x8. Peralatan pengujian yang lengkap dan modern'),
                (37, 'x9. Pengakuan akreditasi laboratorium dari KAN'),
                (38, 'x10. Kemampuan petugas memberi pelayanan informasi'),
                (40, 'x11. Jumlah parameter uji memenuhi standar Permenkes'),
                (41, 'x12. Tampilan Laporan Hasil Pengujian (LHP) mudah dipahami'),
                (42, 'x13. Kemudahan Pelayanan dan Prosedur Pengujian'),
                (43, 'x14. Ketepatan Waktu Penyelesaian Pengujian')
            ]
            jawaban_a = {}
            for baris, teks in aspek_list:
                st.write(f"**{teks}**")
                col1, col2 = st.columns(2)
                with col1: harap = st.slider("Harapan Anda", 1, 4, 4, key=f"h_{baris}")
                with col2: layan = st.slider("Pelayanan Dirasakan", 1, 4, 3, key=f"l_{baris}")
                jawaban_a[baris] = {"harapan": harap, "pelayanan": layan}
                st.write("---")

        with tab2:
            st.write("### Profil & Kebutuhan Pelanggan")
            q1 = st.radio("1. Anda melakukan pengujian untuk kepentingan:", ["Usaha", "Non Usaha"])
            q2 = st.radio("2. Pilihan laboratorium pengujian air Anda:", ["Laboratorium PDAM TKR", "Laboratorium Lain"])
            q2_alasan = st.text_input("Alasan Anda memilih:")
            
            q3 = st.radio("3. Apakah bermaksud melakukan pengujian rutin?", ["Ya", "Tidak"])
            st.write("*Jika Ya, tuliskan kontak yang dapat dihubungi:*")
            c1, c2 = st.columns(2)
            with c1:
                q3_nama = st.text_input("Nama Instansi / Perusahaan:")
                q3_hp = st.text_input("No Telepon / HP:")
            with c2:
                q3_alamat = st.text_area("Alamat:")
                
            q4 = st.radio("4. Apakah parameter pengujian kami memenuhi kebutuhan Anda?", ["Cukup", "Kurang"])
            st.write("*Jika Kurang, parameter apa yang perlu ditambahkan?*")
            param_kurang = {'Amoniak': 70, 'Aluminium': 71, 'Seng': 72, 'Tembaga': 73, 'Detergent': 74, 'Kadmium': 75, 'Chromium Valensi 6': 76, 'Sianida': 77, 'Flourida': 78, 'Phospat': 79}
            pilihan_param = []
            c3, c4 = st.columns(2)
            for i, (param, baris) in enumerate(param_kurang.items()):
                if i % 2 == 0:
                    with c3:
                        if st.checkbox(param): pilihan_param.append(baris)
                else:
                    with c4:
                        if st.checkbox(param): pilihan_param.append(baris)
                        
            q4_lainnya = st.text_input("Parameter Lain-lain:")
            st.write("---")
            q5_saran = st.text_area("5. Mohon saran Anda untuk peningkatan kepuasan pelanggan:")
            
        submit_button = st.form_submit_button("Kirim Kuesioner")

    # PROSES PENYIMPANAN
    if submit_button:
        try:
            wb = openpyxl.load_workbook(TEMPLATE_FILE)
            sheet = wb[NAMA_SHEET]
            
            for baris, skor in jawaban_a.items():
                sheet.cell(row=baris, column=9).value = skor["harapan"]
                sheet.cell(row=baris, column=11).value = skor["pelayanan"]
                
            if q1 == "Usaha": sheet.cell(row=50, column=3).value = "X"
            else: sheet.cell(row=51, column=3).value = "X"
            
            if q2 == "Laboratorium PDAM TKR": sheet.cell(row=54, column=3).value = "X"
            else: sheet.cell(row=55, column=3).value = "X"
            if q2_alasan: sheet.cell(row=56, column=5).value = q2_alasan
                
            if q3 == "Ya": 
                sheet.cell(row=59, column=3).value = "X"
                if q3_nama: sheet.cell(row=62, column=5).value = q3_nama
                if q3_alamat: sheet.cell(row=63, column=5).value = q3_alamat
                if q3_hp: sheet.cell(row=64, column=5).value = q3_hp
            else: sheet.cell(row=60, column=3).value = "X"
                
            if q4 == "Cukup": sheet.cell(row=67, column=3).value = "X"
            else: 
                sheet.cell(row=68, column=3).value = "X"
                for brs_param in pilihan_param: sheet.cell(row=brs_param, column=3).value = "X"
                if q4_lainnya: 
                    sheet.cell(row=80, column=3).value = "X"
                    sheet.cell(row=80, column=5).value = q4_lainnya

            if q5_saran: sheet.cell(row=82, column=3).value = q5_saran

            nama_instansi = q3_nama if q3_nama else "Anonim"
            waktu_sekarang = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            nama_file_unik = f"F13_{nama_instansi}_{datetime.datetime.now().strftime('%H%M%S')}.xlsx"
            path_simpan = os.path.join(FOLDER_HASIL, nama_file_unik)
            
            wb.save(path_simpan)

            df_log = pd.read_csv(LOG_FILE)
            data_baru = pd.DataFrame([{"Waktu Isi": waktu_sekarang, "Nama / Instansi": nama_instansi, "Tujuan Uji": q1}])
            df_log = pd.concat([df_log, data_baru], ignore_index=True)
            df_log.to_csv(LOG_FILE, index=False)

            st.success("Kuesioner Anda berhasil dikirim! Terima kasih atas partisipasinya.")
            
        except Exception as e:
            st.error(f"Terjadi kesalahan sistem: {e}")

# ==========================================
# HALAMAN 2: PANEL ADMIN (UNTUK ANDA)
# ==========================================
elif menu == "Panel Admin":
    st.title("🛡️ Panel Admin")
    st.info("Halaman ini khusus untuk pemantauan kuesioner.")
    
    pwd = st.text_input("Masukkan Password:", type="password")
    
    if pwd == "admin123":
        st.success("Login Berhasil!")
        
        df_log = pd.read_csv(LOG_FILE)
        st.write(f"### Total Responden Saat Ini: {len(df_log)} Orang")
        st.dataframe(df_log, use_container_width=True)
        
        st.write("---")
        st.write("### 📥 Download Data")
        
        if len(os.listdir(FOLDER_HASIL)) > 0:
            zip_buffer = BytesIO()
            with zipfile.ZipFile(zip_buffer, "w") as zip_file:
                for file_name in os.listdir(FOLDER_HASIL):
                    file_path = os.path.join(FOLDER_HASIL, file_name)
                    zip_file.write(file_path, arcname=file_name)
            
            st.download_button(
                label="📦 Download Semua File Kuesioner (.ZIP)",
                data=zip_buffer.getvalue(),
                file_name=f"Rekap_Kuesioner_{datetime.datetime.now().strftime('%Y%m%d')}.zip",
                mime="application/zip",
                help="Download semua file Excel F13 yang sudah diisi pelanggan dalam satu folder ZIP."
            )
            
            with open(LOG_FILE, "rb") as f:
                st.download_button("📄 Download Tabel Log (.CSV)", f, file_name="Log_Responden.csv")
        else:
            st.warning("Belum ada kuesioner yang diisi. Folder masih kosong.")
            
    elif pwd != "":
        st.error("Password Salah!")