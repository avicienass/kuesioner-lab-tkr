import streamlit as st
import openpyxl
from openpyxl.styles import Border, Side, Alignment
from io import BytesIO

# --- Konfigurasi ---
TEMPLATE_FILE = "F 13 - Kuesioner Pelanggan.xlsx"
NAMA_SHEET = "rev 3"

st.set_page_config(page_title="Kuesioner Lab TKR", layout="centered")
st.title("Kuesioner Kepuasan Pelanggan")
st.subheader("Laboratorium Perumdam Tirta Kerta Raharja")

st.write("Terima kasih atas kesediaan Anda menjawab pertanyaan berikut. Berikan nilai **1 (Tidak Baik/Tidak Penting)** hingga **4 (Sangat Baik/Sangat Penting)**.")

tab1, tab2 = st.tabs(["Bagian A: Penilaian Kinerja", "Bagian B: Kebutuhan & Saran"])

with st.form("kuesioner_form"):
    
    # === TAB 1: PENILAIAN (14 Aspek) ===
    with tab1:
        st.write("### Penilaian Kinerja Laboratorium")
        st.info("Nilai 1-4 untuk tingkat Harapan (Kepentingan) & Pelayanan (Kenyataan).")
        
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
            with col1:
                harap = st.slider("Harapan Anda", 1, 4, 4, key=f"h_{baris}")
            with col2:
                layan = st.slider("Pelayanan Dirasakan", 1, 4, 3, key=f"l_{baris}")
            jawaban_a[baris] = {"harapan": harap, "pelayanan": layan}
            st.write("---")

    # === TAB 2: BAGIAN B & SARAN ===
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
        
        param_kurang = {
            'Amoniak': 70, 'Aluminium': 71, 'Seng': 72, 'Tembaga': 73,
            'Detergent': 74, 'Kadmium': 75, 'Chromium Valensi 6': 76,
            'Sianida': 77, 'Flourida': 78, 'Phospat': 79
        }
        
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
        st.write("### Saran Peningkatan")
        q5_saran = st.text_area("5. Mohon saran Anda untuk peningkatan kepuasan pelanggan:")
        
    submit_button = st.form_submit_button("Simpan & Cetak Kuesioner")

# --- PROSES MEMASUKKAN KE EXCEL ---
if submit_button:
    try:
        wb = openpyxl.load_workbook(TEMPLATE_FILE)
        sheet = wb[NAMA_SHEET]
        
        # Pengaturan gaya (style) untuk kotak "X"
        garis_tipis = Side(border_style="thin", color="000000")
        kotak_border = Border(top=garis_tipis, left=garis_tipis, right=garis_tipis, bottom=garis_tipis)
        posisi_tengah = Alignment(horizontal="center", vertical="center")
        
        # Fungsi pembantu untuk membuat kotak X yang rapi
        def beri_silang_kotak(baris, kolom):
            sel = sheet.cell(row=baris, column=kolom)
            sel.value = "X"
            sel.border = kotak_border
            sel.alignment = posisi_tengah

        # 1. Mengisi Bagian A
        for baris, skor in jawaban_a.items():
            # Untuk bagian A kita hanya mengatur alignment tengah agar rapi, tanpa kotak border
            sel_harap = sheet.cell(row=baris, column=9)
            sel_harap.value = skor["harapan"]
            sel_harap.alignment = posisi_tengah
            
            sel_layan = sheet.cell(row=baris, column=11)
            sel_layan.value = skor["pelayanan"]
            sel_layan.alignment = posisi_tengah
            
        # 2. Mengisi Bagian B dengan "X" di dalam kotak
        if q1 == "Usaha": beri_silang_kotak(50, 3)
        else: beri_silang_kotak(51, 3)
        
        if q2 == "Laboratorium PDAM TKR": beri_silang_kotak(54, 3)
        else: beri_silang_kotak(55, 3)
        
        if q2_alasan: sheet.cell(row=56, column=5).value = q2_alasan
            
        if q3 == "Ya": 
            beri_silang_kotak(59, 3)
            if q3_nama: sheet.cell(row=62, column=5).value = q3_nama
            if q3_alamat: sheet.cell(row=63, column=5).value = q3_alamat
            if q3_hp: sheet.cell(row=64, column=5).value = q3_hp
        else: 
            beri_silang_kotak(60, 3)
            
        if q4 == "Cukup": 
            beri_silang_kotak(67, 3)
        else: 
            beri_silang_kotak(68, 3)
            for brs_param in pilihan_param:
                beri_silang_kotak(brs_param, 3)
            if q4_lainnya: 
                beri_silang_kotak(80, 3)
                sheet.cell(row=80, column=5).value = q4_lainnya

        # 3. Mengisi Saran
        if q5_saran:
            sheet.cell(row=82, column=3).value = q5_saran

        output = BytesIO()
        wb.save(output)
        output.seek(0)
        
        st.success("Kuesioner berhasil diproses dengan tanda X di dalam kotak yang rapi!")
        
        nama_download = q3_nama if q3_nama else "Pelanggan"
        
        st.download_button(
            label="📥 Download File F13 (Terisi Otomatis)",
            data=output,
            file_name=f"F13_Terisi_{nama_download}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
        
    except Exception as e:
        st.error(f"Gagal memproses: {e}")