import datetime
import io
import os
from docx import Document
from docx.shared import RGBColor
import streamlit as st

# Konfigurasi Halaman
st.set_page_config(
    page_title="Formulir Penyerahan Skripsi SPI UIN Jakarta",
    page_icon="📄",
    layout="centered"
)

# Pemetaan Nama Bulan Bahasa Indonesia
BULAN_ID = {
    1: "Januari", 2: "Februari", 3: "Maret", 4: "April",
    5: "Mei", 6: "Juni", 7: "Juli", 8: "Agustus",
    9: "September", 10: "Oktober", 11: "November", 12: "Desember"
}

def format_tanggal_indo(d: datetime.date) -> str:
    if not d:
        return ""
    return f"{d.day} {BULAN_ID.get(d.month, '')} {d.year}"

def generate_document(template_path: str, data: dict) -> io.BytesIO:
    doc = Document(template_path)

    # 1. Nama Mahasiswa (P1)
    doc.paragraphs[1].runs[-1].text = f"\t: {data['nama']}"

    # 2. NIM (P2)
    doc.paragraphs[2].runs[-1].text = f"\t: {data['nim']}"

    # 3. Program Studi (P3)
    if data.get("prodi"):
        doc.paragraphs[3].runs[-1].text = data["prodi"]

    # 4. Tanggal Sidang (P4)
    # Ubah run terakhir agar tidak merah dan masukkan tanggal sidang
    doc.paragraphs[4].runs[-1].text = f": {data['tanggal_sidang']}"
    doc.paragraphs[4].runs[-1].font.color.rgb = RGBColor(0, 0, 0)

    # 5. Judul Skripsi (P5)
    doc.paragraphs[5].runs[-1].text = f"\t: {data['judul']}"

    # 6. Tempat, Tanggal Dokumen (P9 - sebelumnya berwarna merah di template asli)
    tempat_tgl = f"{data['tempat']}, {data['tanggal_penyerahan']}"
    doc.paragraphs[9].runs[0].text = tempat_tgl
    doc.paragraphs[9].runs[0].font.color.rgb = RGBColor(0, 0, 0)
    for r in doc.paragraphs[9].runs[1:]:
        r.text = ""

    # 7. Data Dosen Pembimbing & Penguji (Tabel baris 2, 3, 4)
    table = doc.tables[0]

    def update_dosen(cell, nama, nip):
        if nama.strip():
            cell.paragraphs[0].text = nama.strip()
        if nip.strip():
            clean_nip = nip.strip()
            if not clean_nip.lower().startswith(("nip", "nidn")):
                clean_nip = f"NIP/NIDN. {clean_nip}"
            cell.paragraphs[1].text = clean_nip

    if data.get("pembimbing_nama"):
        update_dosen(table.rows[2].cells[1], data["pembimbing_nama"], data.get("pembimbing_nip", ""))
    if data.get("penguji1_nama"):
        update_dosen(table.rows[3].cells[1], data["penguji1_nama"], data.get("penguji1_nip", ""))
    if data.get("penguji2_nama"):
        update_dosen(table.rows[4].cells[1], data["penguji2_nama"], data.get("penguji2_nip", ""))

    output_stream = io.BytesIO()
    doc.save(output_stream)
    output_stream.seek(0)
    return output_stream

# UI Header
st.title("📄 Generator Tanda Penyerahan Skripsi")
st.subheader("Program Studi Sejarah dan Peradaban Islam (SPI)")
st.caption("Fakultas Adab dan Humaniora — UIN Syarif Hidayatullah Jakarta")

st.info("💡 Isi data di bawah ini untuk membuat dokumen formulir resmi secara otomatis tanpa perlu mengubah format manual di Word.")

# Form Input
with st.form("form_skripsi"):
    col1, col2 = st.columns(2)
    with col1:
        nama = st.text_input("Nama Lengkap", placeholder="Contoh: Syarif Hidayatullah")
    with col2:
        nim = st.text_input("NIM", placeholder="Contoh: 11200210000088")

    prodi = st.text_input("Program Studi", value="Sejarah dan Peradaban Islam")

    tgl_sidang_raw = st.date_input("Tanggal Sidang Munaqasyah", value=datetime.date.today())

    judul = st.text_area("Judul Skripsi", placeholder="Tuliskan judul skripsi lengkap sesuai naskah final...", height=100)

    st.markdown("---")
    st.markdown("##### Tempat & Tanggal Penandatanganan")
    col_t1, col_t2 = st.columns(2)
    with col_t1:
        tempat = st.text_input("Tempat Surat", value="Jakarta")
    with col_t2:
        tgl_penyerahan_raw = st.date_input("Tanggal Dokumen", value=datetime.date.today())

    st.markdown("---")
    with st.expander("🧑‍🏫 Data Dosen Pembimbing & Penguji (Opsional)"):
        st.caption("Biarkan kosong jika ingin mengikuti format bawaan template (Nama Dosen / NIP).")
        
        st.markdown("**Dosen Pembimbing Skripsi**")
        col_p1, col_p2 = st.columns(2)
        with col_p1:
            pembimbing_nama = st.text_input("Nama & Gelar Pembimbing", placeholder="Contoh: Dr. Zakiya Darojat, M.A.")
        with col_p2:
            pembimbing_nip = st.text_input("NIP / NIDN Pembimbing", placeholder="Contoh: 197405302005012006")

        st.markdown("**Dosen Penguji 1**")
        col_u1, col_u2 = st.columns(2)
        with col_u1:
            penguji1_nama = st.text_input("Nama & Gelar Penguji 1", placeholder="Nama Dosen Penguji 1")
        with col_u2:
            penguji1_nip = st.text_input("NIP / NIDN Penguji 1", placeholder="NIP/NIDN")

        st.markdown("**Dosen Penguji 2**")
        col_u3, col_u4 = st.columns(2)
        with col_u3:
            penguji2_nama = st.text_input("Nama & Gelar Penguji 2", placeholder="Nama Dosen Penguji 2")
        with col_u4:
            penguji2_nip = st.text_input("NIP / NIDN Penguji 2", placeholder="NIP/NIDN")

    submitted = st.form_submit_button("🚀 Buat Dokumen (.docx)", use_container_width=True)

if submitted:
    if not nama.strip():
        st.error("⚠️ Nama Lengkap wajib diisi!")
    elif not nim.strip():
        st.error("⚠️ NIM wajib diisi!")
    elif not judul.strip():
        st.error("⚠️ Judul Skripsi wajib diisi!")
    else:
        template_file = os.path.join(os.path.dirname(__file__), "template.docx")
        if not os.path.exists(template_file):
            st.error("❌ File template.docx tidak ditemukan di direktori aplikasi!")
        else:
            payload = {
                "nama": nama.strip(),
                "nim": nim.strip(),
                "prodi": prodi.strip(),
                "tanggal_sidang": format_tanggal_indo(tgl_sidang_raw),
                "judul": judul.strip(),
                "tempat": tempat.strip() or "Jakarta",
                "tanggal_penyerahan": format_tanggal_indo(tgl_penyerahan_raw),
                "pembimbing_nama": pembimbing_nama,
                "pembimbing_nip": pembimbing_nip,
                "penguji1_nama": penguji1_nama,
                "penguji1_nip": penguji1_nip,
                "penguji2_nama": penguji2_nama,
                "penguji2_nip": penguji2_nip,
            }

            doc_bytes = generate_document(template_file, payload)
            clean_nim = "".join(c for c in nim if c.isalnum())
            filename = f"Tanda_Bukti_Penyerahan_Skripsi_{clean_nim}.docx"

            st.success("✅ Dokumen berhasil dibuat dengan rapi!")
            st.download_button(
                label="📥 Unduh Dokumen Word (.docx)",
                data=doc_bytes,
                file_name=filename,
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                use_container_width=True
            )
