import datetime
import io
import os
import re
import shutil
import subprocess
import tempfile
from docx import Document
from docx.shared import Mm, Pt, RGBColor
from docx.oxml import parse_xml
from docx.oxml.ns import nsdecls
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

def format_nip_nidn(raw_val: str) -> str:
    """
    Mendeteksi otomatis NIP vs NIDN:
    - 18 digit: NIP (PNS/ASN)
    - 10 digit: NIDN (Dosen Tetap/Non-PNS)
    - Jika user sudah menulis 'NIP' atau 'NIDN', gunakan input apa adanya.
    """
    val = raw_val.strip()
    if not val:
        return ""
    if val.upper().startswith(("NIP", "NIDN")):
        return val
    
    digits = re.sub(r"\D", "", val)
    if len(digits) == 18:
        return f"NIP. {val}"
    elif len(digits) == 10:
        return f"NIDN. {val}"
    else:
        return f"NIP/NIDN. {val}"

def generate_document(template_path: str, data: dict) -> bytes:
    doc = Document(template_path)

    # 1. Atur margin agar pas 1 lembar A4
    sec = doc.sections[0]
    sec.top_margin = Mm(15)
    sec.bottom_margin = Mm(12)
    sec.left_margin = Mm(20)
    sec.right_margin = Mm(20)

    # 2. Nama Mahasiswa (P1)
    doc.paragraphs[1].runs[-1].text = f"\t: {data['nama']}"

    # 3. NIM (P2)
    doc.paragraphs[2].runs[-1].text = f"\t: {data['nim']}"

    # 4. Program Studi (P3)
    if data.get("prodi"):
        doc.paragraphs[3].runs[-1].text = data["prodi"]

    # 5. Tanggal Sidang (P4)
    doc.paragraphs[4].runs[-1].text = f": {data['tanggal_sidang']}"
    doc.paragraphs[4].runs[-1].font.color.rgb = RGBColor(0, 0, 0)

    # 6. Judul Skripsi (P5)
    doc.paragraphs[5].runs[-1].text = f"\t: {data['judul']}"

    # 7. Padding 1 baris antara judul skripsi dan tabel (pertahankan P6, hapus P7)
    p7 = doc.paragraphs[7]._p
    p7.getparent().remove(p7)
    p6 = doc.paragraphs[6]
    p6.paragraph_format.space_before = Pt(0)
    p6.paragraph_format.space_after = Pt(0)
    p6.paragraph_format.line_spacing = 1.0

    # 8. Optimalisasi ukuran baris dan teks tabel agar hemat ruang dan presisi
    table = doc.tables[0]
    xml_ns = nsdecls("w")
    for r in table.rows:
        trPr = r._tr.get_or_add_trPr()
        for th in trPr.xpath("./w:trHeight"):
            trPr.remove(th)
        new_th = parse_xml(f'<w:trHeight {xml_ns} w:val="380" w:hRule="atLeast"/>')
        trPr.append(new_th)
        for c in r.cells:
            for p in c.paragraphs:
                p.paragraph_format.space_before = Pt(1)
                p.paragraph_format.space_after = Pt(1)
                p.paragraph_format.line_spacing = 1.0
                for run in p.runs:
                    run.font.size = Pt(10.5)

    # 9. Data Dosen Pembimbing & Penguji (Tabel baris 2, 3, 4)
    def update_dosen(cell, nama, nip_raw):
        if nama.strip():
            cell.paragraphs[0].text = nama.strip()
            for r in cell.paragraphs[0].runs:
                r.font.size = Pt(10.5)
        if nip_raw.strip():
            cell.paragraphs[1].text = format_nip_nidn(nip_raw)
            for r in cell.paragraphs[1].runs:
                r.font.size = Pt(10.5)

    if data.get("pembimbing_nama"):
        update_dosen(table.rows[2].cells[1], data["pembimbing_nama"], data.get("pembimbing_nip", ""))
    if data.get("penguji1_nama"):
        update_dosen(table.rows[3].cells[1], data["penguji1_nama"], data.get("penguji1_nip", ""))
    if data.get("penguji2_nama"):
        update_dosen(table.rows[4].cells[1], data["penguji2_nama"], data.get("penguji2_nip", ""))

    # 10. Hapus tab kosong sebelum tanggal
    p_tabs = doc.paragraphs[7]._p
    p_tabs.getparent().remove(p_tabs)

    # 11. Tempat & Tanggal Dokumen (P7 setelah penghapusan p_tabs)
    p_jakarta = doc.paragraphs[7]
    tempat_tgl = f"{data['tempat']}, {data['tanggal_penyerahan']}"
    p_jakarta.runs[0].text = tempat_tgl
    p_jakarta.runs[0].font.color.rgb = RGBColor(0, 0, 0)
    for r in p_jakarta.runs[1:]:
        r.text = ""

    # 12. Hapus tab kosong setelah tanggal
    p_tabs2 = doc.paragraphs[8]._p
    p_tabs2.getparent().remove(p_tabs2)

    # 13. Hapus 2 baris kosong manual tanda tangan agar tidak meluber ke lembar ke-2
    p_sig1 = doc.paragraphs[10]._p
    p_sig1.getparent().remove(p_sig1)
    p_sig2 = doc.paragraphs[10]._p
    p_sig2.getparent().remove(p_sig2)

    # Beri spasi tanda tangan vertikal yang pas (40 pt) langsung pada nama Kabag TU
    p_iwan = doc.paragraphs[10]
    p_iwan.paragraph_format.space_before = Pt(40)

    output_stream = io.BytesIO()
    doc.save(output_stream)
    return output_stream.getvalue()

def convert_docx_to_pdf(docx_bytes: bytes) -> bytes | None:
    soffice_bin = shutil.which("soffice") or shutil.which("libreoffice")
    if not soffice_bin:
        return None
    with tempfile.TemporaryDirectory() as tmpdir:
        input_path = os.path.join(tmpdir, "input.docx")
        with open(input_path, "wb") as f:
            f.write(docx_bytes)
        cmd = [soffice_bin, "--headless", "--convert-to", "pdf", "--outdir", tmpdir, input_path]
        try:
            res = subprocess.run(cmd, stdout=subprocess.PIPE, stderr=subprocess.PIPE, timeout=30)
            output_pdf = os.path.join(tmpdir, "input.pdf")
            if os.path.exists(output_pdf):
                with open(output_pdf, "rb") as pf:
                    return pf.read()
        except Exception:
            return None
    return None

# UI Header
st.title("📄 Generator Tanda Penyerahan Skripsi")
st.subheader("Program Studi Sejarah dan Peradaban Islam (SPI)")
st.caption("Fakultas Adab dan Humaniora — UIN Syarif Hidayatullah Jakarta")

st.info("💡 Hasil dokumen dijamin pas **1 lembar A4** (tidak meluber ke halaman 2). Format tersedia dalam **PDF (siap cetak)** dan **Word (.docx)**.")

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
    st.markdown("##### Tempat & Tanggal Penandatanganan Surat")
    col_t1, col_t2 = st.columns(2)
    with col_t1:
        tempat = st.text_input("Tempat Surat", value="Jakarta")
    with col_t2:
        tgl_penyerahan_raw = st.date_input("Tanggal Dokumen", value=datetime.date.today())

    st.markdown("---")
    with st.expander("🧑‍🏫 Data Dosen Pembimbing & Penguji (Opsional)"):
        st.caption("Ketik angka nomornya saja. Sistem otomatis mendeteksi: **18 digit = NIP**, **10 digit = NIDN**.")
        
        st.markdown("**Dosen Pembimbing Skripsi**")
        col_p1, col_p2 = st.columns(2)
        with col_p1:
            pembimbing_nama = st.text_input("Nama & Gelar Pembimbing", placeholder="Contoh: Dr. Zakiya Darojat, M.A.")
        with col_p2:
            pembimbing_nip = st.text_input("Nomor NIP / NIDN Pembimbing", placeholder="18 digit (NIP) atau 10 digit (NIDN)")

        st.markdown("**Dosen Penguji 1**")
        col_u1, col_u2 = st.columns(2)
        with col_u1:
            penguji1_nama = st.text_input("Nama & Gelar Penguji 1", placeholder="Nama Dosen Penguji 1")
        with col_u2:
            penguji1_nip = st.text_input("Nomor NIP / NIDN Penguji 1", placeholder="18 digit (NIP) atau 10 digit (NIDN)")

        st.markdown("**Dosen Penguji 2**")
        col_u3, col_u4 = st.columns(2)
        with col_u3:
            penguji2_nama = st.text_input("Nama & Gelar Penguji 2", placeholder="Nama Dosen Penguji 2")
        with col_u4:
            penguji2_nip = st.text_input("Nomor NIP / NIDN Penguji 2", placeholder="18 digit (NIP) atau 10 digit (NIDN)")

    submitted = st.form_submit_button("🚀 Proses & Buat Dokumen (1 Lembar)", use_container_width=True)

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

            with st.spinner("Sedang memproses dokumen agar pas 1 lembar..."):
                docx_bytes = generate_document(template_file, payload)
                pdf_bytes = convert_docx_to_pdf(docx_bytes)
                clean_nim = "".join(c for c in nim if c.isalnum())

            st.success("✅ Dokumen berhasil dibuat tepat 1 lembar A4!")
            st.markdown("##### Pilih format unduhan:")

            col_dl1, col_dl2 = st.columns(2)
            with col_dl1:
                if pdf_bytes:
                    st.download_button(
                        label="📄 Unduh PDF (Pas 1 Lembar)",
                        data=pdf_bytes,
                        file_name=f"Tanda_Bukti_Penyerahan_Skripsi_{clean_nim}.pdf",
                        mime="application/pdf",
                        use_container_width=True
                    )
                else:
                    st.warning("Konversi PDF tidak tersedia.")

            with col_dl2:
                st.download_button(
                    label="📝 Unduh Word (.docx)",
                    data=docx_bytes,
                    file_name=f"Tanda_Bukti_Penyerahan_Skripsi_{clean_nim}.docx",
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                    use_container_width=True
                )
