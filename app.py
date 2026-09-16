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
    page_title="Layanan Berkas Skripsi SPI UIN Jakarta",
    page_icon="🎓",
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

# ==========================================
# FUNGSI DOKUMEN 1: PENYERAHAN SKRIPSI
# ==========================================
def generate_penyerahan_document(template_path: str, data: dict) -> bytes:
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

# ==========================================
# FUNGSI DOKUMEN 2: PENDAFTARAN UJIAN SKRIPSI
# ==========================================
def generate_ujian_document(template_path: str, data: dict) -> bytes:
    doc = Document(template_path)

    # 1. Nama Lengkap (P1)
    doc.paragraphs[1].runs[-1].text = f"\t: {data['nama']}"

    # 2. NIM (P2)
    doc.paragraphs[2].runs[-1].text = f"\t: {data['nim']}"

    # 3. Program Studi (P3)
    doc.paragraphs[3].runs[-1].text = f"\t: {data['prodi']}"

    # 4. Tempat/Tanggal Lahir (P4)
    doc.paragraphs[4].runs[-1].text = f"\t: {data['ttl']}"

    # 5. Asal SLTA (P5)
    doc.paragraphs[5].runs[-1].text = f"\t: {data['asal_slta']}"

    # 6. Alamat Sekarang (P6)
    doc.paragraphs[6].runs[-1].text = f"\t: {data['alamat']}"

    # 7. IPK Sementara (P7)
    doc.paragraphs[7].runs[-1].text = f"\t: {data['ipk']}"

    # 8. Judul Skripsi (P8)
    doc.paragraphs[8].runs[-1].text = f"\t: {data['judul']}"

    # 9. No. Telpon/Email (P9)
    doc.paragraphs[9].runs[-1].text = f": {data['kontak']}"

    # 10. Kalimat Permohonan (P11)
    doc.paragraphs[11].runs[0].text = f"Dengan ini mengajukan permohonan ujian skripsi pada semester {data['semester']} Tahun Akademik {data['tahun_akademik']}"

    # 11. Tempat & Tanggal Dokumen (P13)
    doc.paragraphs[13].runs[0].text = f"{data['tempat']}, {data['tanggal']}"
    for r in doc.paragraphs[13].runs[1:]:
        r.text = ""

    # 12. Nama & NIM Mahasiswa (P16, P17)
    doc.paragraphs[16].runs[0].text = data["nama"]
    doc.paragraphs[17].runs[0].text = f"NIM : {data['nim']}"

    # 13. Hapus baris Catatan & Tanda Bintang di bagian bawah jika ada
    for i in range(len(doc.paragraphs) - 1, -1, -1):
        txt = doc.paragraphs[i].text.strip()
        if txt.startswith("Catatan") or "Tanda Bintang" in txt:
            p_el = doc.paragraphs[i]._p
            p_el.getparent().remove(p_el)

    output_stream = io.BytesIO()
    doc.save(output_stream)
    return output_stream.getvalue()

# ==========================================
# FUNGSI DOKUMEN 3: PENDAFTARAN SIDANG SKRIPSI (MAP KUNING)
# ==========================================
def generate_sidang_document(template_path: str, data: dict) -> bytes:
    doc = Document(template_path)

    # 1. Nama Mahasiswa (P0)
    doc.paragraphs[0].runs[-1].text = f"\t: {data['nama']}"

    # 2. NIM (P1)
    doc.paragraphs[1].runs[-1].text = f"\t: {data['nim']}"

    # 3. Program Studi (P2)
    doc.paragraphs[2].runs[-1].text = f"\t: {data['prodi']}"

    # 4. Tanggal & Tempat Dokumen (P19)
    doc.paragraphs[19].runs[0].text = f"\t{data['tempat']}, {data['tanggal']}"
    for r in doc.paragraphs[19].runs[1:]:
        r.text = ""

    # 5. Perbaiki koma ganda pada nama Kabag TU
    doc.paragraphs[24].runs[-1].text = "Iwan Kurniawan, S.Pd., M.Si."

    output_stream = io.BytesIO()
    doc.save(output_stream)
    return output_stream.getvalue()

# Header Utama Aplikasi
st.title("🎓 Portal Layanan Berkas Skripsi")
st.subheader("Program Studi Sejarah dan Peradaban Islam (SPI)")
st.caption("Fakultas Adab dan Humaniora — UIN Syarif Hidayatullah Jakarta")

# Tab Pilihan Berkas (3 Formulir Lengkap)
tab1, tab2, tab3 = st.tabs([
    "📄 Tanda Bukti Penyerahan Skripsi",
    "📝 Formulir Pendaftaran Ujian Skripsi",
    "📋 Formulir Pendaftaran Sidang (Persyaratan)"
])

# ==========================================
# TAB 1: FORM PENYERAHAN SKRIPSI
# ==========================================
with tab1:
    st.info("💡 Formulir tanda bukti penyerahan skripsi pasca munaqasyah ke pihak Fakultas, Prodi, dan Perpustakaan.")
    
    with st.form("form_penyerahan"):
        col1, col2 = st.columns(2)
        with col1:
            p_nama = st.text_input("Nama Lengkap Mahasiswa", placeholder="Contoh: Syarif Hidayatullah", key="p_nama")
        with col2:
            p_nim = st.text_input("NIM", placeholder="Contoh: 11200210000088", key="p_nim")

        p_prodi = st.text_input("Program Studi", value="Sejarah dan Peradaban Islam", key="p_prodi")
        p_tgl_sidang = st.date_input("Tanggal Sidang Munaqasyah", value=datetime.date.today(), key="p_sidang")
        p_judul = st.text_area("Judul Skripsi", placeholder="Tuliskan judul skripsi naskah final...", height=90, key="p_judul")

        st.markdown("##### Tempat & Tanggal Penandatanganan Surat")
        col_pt1, col_pt2 = st.columns(2)
        with col_pt1:
            p_tempat = st.text_input("Tempat Surat", value="Jakarta", key="p_tempat")
        with col_pt2:
            p_tgl_penyerahan = st.date_input("Tanggal Surat", value=datetime.date.today(), key="p_tgl")

        with st.expander("🧑‍🏫 Data Dosen Pembimbing & Penguji (Opsional)"):
            st.caption("Ketik angka nomornya saja. Sistem otomatis mendeteksi: **18 digit = NIP**, **10 digit = NIDN**.")
            st.markdown("**Dosen Pembimbing Skripsi**")
            col_pp1, col_pp2 = st.columns(2)
            with col_pp1:
                p_pembimbing_nama = st.text_input("Nama & Gelar Pembimbing", placeholder="Nama Pembimbing", key="p_pnama")
            with col_pp2:
                p_pembimbing_nip = st.text_input("Nomor NIP / NIDN Pembimbing", placeholder="Nomor NIP/NIDN", key="p_pnip")

            st.markdown("**Dosen Penguji 1**")
            col_pu1, col_pu2 = st.columns(2)
            with col_pu1:
                p_penguji1_nama = st.text_input("Nama & Gelar Penguji 1", placeholder="Nama Penguji 1", key="p_u1nama")
            with col_pu2:
                p_penguji1_nip = st.text_input("Nomor NIP / NIDN Penguji 1", placeholder="Nomor NIP/NIDN", key="p_u1nip")

            st.markdown("**Dosen Penguji 2**")
            col_pu3, col_pu4 = st.columns(2)
            with col_pu3:
                p_penguji2_nama = st.text_input("Nama & Gelar Penguji 2", placeholder="Nama Penguji 2", key="p_u2nama")
            with col_pu4:
                p_penguji2_nip = st.text_input("Nomor NIP / NIDN Penguji 2", placeholder="Nomor NIP/NIDN", key="p_u2nip")

        p_submitted = st.form_submit_button("🚀 Buat Dokumen Penyerahan (1 Lembar)", use_container_width=True)

    if p_submitted:
        if not p_nama.strip():
            st.error("⚠️ Nama Lengkap wajib diisi!")
        elif not p_nim.strip():
            st.error("⚠️ NIM wajib diisi!")
        elif not p_judul.strip():
            st.error("⚠️ Judul Skripsi wajib diisi!")
        else:
            tpl_penyerahan = os.path.join(os.path.dirname(__file__), "template.docx")
            if not os.path.exists(tpl_penyerahan):
                st.error("❌ File template.docx tidak ditemukan di direktori aplikasi!")
            else:
                payload_p = {
                    "nama": p_nama.strip(),
                    "nim": p_nim.strip(),
                    "prodi": p_prodi.strip(),
                    "tanggal_sidang": format_tanggal_indo(p_tgl_sidang),
                    "judul": p_judul.strip(),
                    "tempat": p_tempat.strip() or "Jakarta",
                    "tanggal_penyerahan": format_tanggal_indo(p_tgl_penyerahan),
                    "pembimbing_nama": p_pembimbing_nama,
                    "pembimbing_nip": p_pembimbing_nip,
                    "penguji1_nama": p_penguji1_nama,
                    "penguji1_nip": p_penguji1_nip,
                    "penguji2_nama": p_penguji2_nama,
                    "penguji2_nip": p_penguji2_nip,
                }

                with st.spinner("Sedang memproses dokumen agar pas 1 lembar..."):
                    docx_bytes_p = generate_penyerahan_document(tpl_penyerahan, payload_p)
                    pdf_bytes_p = convert_docx_to_pdf(docx_bytes_p)
                    clean_nim_p = "".join(c for c in p_nim if c.isalnum())

                st.success("✅ Dokumen berhasil dibuat tepat 1 lembar A4!")
                st.markdown("##### Pilih format unduhan:")
                col_pdl1, col_pdl2 = st.columns(2)
                with col_pdl1:
                    if pdf_bytes_p:
                        st.download_button(
                            label="📄 Unduh PDF (Pas 1 Lembar)",
                            data=pdf_bytes_p,
                            file_name=f"Tanda_Bukti_Penyerahan_Skripsi_{clean_nim_p}.pdf",
                            mime="application/pdf",
                            use_container_width=True
                        )
                    else:
                        st.warning("Konversi PDF tidak tersedia.")
                with col_pdl2:
                    st.download_button(
                        label="📝 Unduh Word (.docx)",
                        data=docx_bytes_p,
                        file_name=f"Tanda_Bukti_Penyerahan_Skripsi_{clean_nim_p}.docx",
                        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                        use_container_width=True
                    )

# ==========================================
# TAB 2: FORM PENDAFTARAN UJIAN SKRIPSI
# ==========================================
with tab2:
    st.info("💡 Formulir permohonan pendaftaran ujian skripsi munaqasyah untuk diajukan ke Tata Usaha/Fakultas.")
    
    with st.form("form_ujian"):
        col_u1, col_u2 = st.columns(2)
        with col_u1:
            u_nama = st.text_input("Nama Lengkap", placeholder="Contoh: Syarif Hidayatullah", key="u_nama")
        with col_u2:
            u_nim = st.text_input("NIM", placeholder="Contoh: 11200210000088", key="u_nim")

        u_prodi = st.text_input("Program Studi", value="Sejarah dan Peradaban Islam", key="u_prodi")

        col_ttl1, col_ttl2 = st.columns(2)
        with col_ttl1:
            u_kota_lahir = st.text_input("Tempat Lahir", placeholder="Contoh: Tangerang", key="u_kota")
        with col_ttl2:
            u_tgl_lahir = st.date_input("Tanggal Lahir", value=datetime.date(2002, 1, 1), key="u_tgl_lahir")

        col_slta, col_ipk = st.columns(2)
        with col_slta:
            u_asal_slta = st.text_input("Asal SLTA (SMA/SMK/MA)", placeholder="Contoh: MAN 1 Tangerang", key="u_slta")
        with col_ipk:
            u_ipk = st.text_input("IPK Sementara", placeholder="Contoh: 3.75", key="u_ipk")

        u_alamat = st.text_area("Alamat Sekarang", placeholder="Alamat lengkap domisili saat ini...", height=80, key="u_alamat")
        u_judul = st.text_area("Judul Skripsi", placeholder="Tuliskan judul skripsi lengkap...", height=90, key="u_judul")

        col_k1, col_k2 = st.columns(2)
        with col_k1:
            u_telp = st.text_input("Nomor Telepon / WhatsApp", placeholder="Contoh: 081234567890", key="u_telp")
        with col_k2:
            u_email = st.text_input("Alamat Email", placeholder="Contoh: syarif@gmail.com", key="u_email")

        col_sem1, col_sem2 = st.columns(2)
        with col_sem1:
            u_semester = st.selectbox("Semester Permohonan", ["Ganjil", "Genap"], key="u_sem")
        with col_sem2:
            u_thn_akademik = st.text_input("Tahun Akademik", value="2025/2026", key="u_thn")

        st.markdown("##### Tempat & Tanggal Penandatanganan Surat")
        col_ut1, col_ut2 = st.columns(2)
        with col_ut1:
            u_tempat = st.text_input("Tempat Surat", value="Jakarta", key="u_tempat")
        with col_ut2:
            u_tgl_surat = st.date_input("Tanggal Dokumen", value=datetime.date.today(), key="u_tgl_surat")

        u_submitted = st.form_submit_button("🚀 Buat Dokumen Pendaftaran Ujian (1 Lembar)", use_container_width=True)

    if u_submitted:
        if not u_nama.strip():
            st.error("⚠️ Nama Lengkap wajib diisi!")
        elif not u_nim.strip():
            st.error("⚠️ NIM wajib diisi!")
        elif not u_judul.strip():
            st.error("⚠️ Judul Skripsi wajib diisi!")
        else:
            tpl_ujian = os.path.join(os.path.dirname(__file__), "template_ujian.docx")
            if not os.path.exists(tpl_ujian):
                st.error("❌ File template_ujian.docx tidak ditemukan di direktori aplikasi!")
            else:
                # Format kontak gabungan
                kontak_parts = []
                if u_telp.strip():
                    kontak_parts.append(u_telp.strip())
                if u_email.strip():
                    kontak_parts.append(u_email.strip())
                kontak_str = " / ".join(kontak_parts) if kontak_parts else "-"

                ttl_str = f"{u_kota_lahir.strip()}, {format_tanggal_indo(u_tgl_lahir)}" if u_kota_lahir.strip() else format_tanggal_indo(u_tgl_lahir)

                payload_u = {
                    "nama": u_nama.strip(),
                    "nim": u_nim.strip(),
                    "prodi": u_prodi.strip(),
                    "ttl": ttl_str,
                    "asal_slta": u_asal_slta.strip() or "-",
                    "alamat": u_alamat.strip() or "-",
                    "ipk": u_ipk.strip() or "-",
                    "judul": u_judul.strip(),
                    "kontak": kontak_str,
                    "semester": u_semester,
                    "tahun_akademik": u_thn_akademik.strip() or "2025/2026",
                    "tempat": u_tempat.strip() or "Jakarta",
                    "tanggal": format_tanggal_indo(u_tgl_surat)
                }

                with st.spinner("Sedang memproses dokumen formulir ujian..."):
                    docx_bytes_u = generate_ujian_document(tpl_ujian, payload_u)
                    pdf_bytes_u = convert_docx_to_pdf(docx_bytes_u)
                    clean_nim_u = "".join(c for c in u_nim if c.isalnum())

                st.success("✅ Formulir Pendaftaran Ujian berhasil dibuat tepat 1 lembar A4!")
                st.markdown("##### Pilih format unduhan:")
                col_udl1, col_udl2 = st.columns(2)
                with col_udl1:
                    if pdf_bytes_u:
                        st.download_button(
                            label="📄 Unduh PDF (Pas 1 Lembar)",
                            data=pdf_bytes_u,
                            file_name=f"Pendaftaran_Ujian_Skripsi_{clean_nim_u}.pdf",
                            mime="application/pdf",
                            use_container_width=True
                        )
                    else:
                        st.warning("Konversi PDF tidak tersedia.")
                with col_udl2:
                    st.download_button(
                        label="📝 Unduh Word (.docx)",
                        data=docx_bytes_u,
                        file_name=f"Pendaftaran_Ujian_Skripsi_{clean_nim_u}.docx",
                        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                        use_container_width=True
                    )

# ==========================================
# TAB 3: FORMULIR PENDAFTARAN SIDANG (PERSYARATAN MAP KUNING)
# ==========================================
with tab3:
    st.info("💡 Formulir pendaftaran sidang skripsi yang memuat 11 berkas checklist persyaratan untuk dimasukkan ke dalam map kuning.")
    
    with st.form("form_sidang"):
        col_s1, col_s2 = st.columns(2)
        with col_s1:
            s_nama = st.text_input("Nama Lengkap", placeholder="Contoh: Syarif Hidayatullah", key="s_nama")
        with col_s2:
            s_nim = st.text_input("NIM", placeholder="Contoh: 11200210000088", key="s_nim")

        s_prodi = st.text_input("Program Studi", value="Sejarah dan Peradaban Islam", key="s_prodi")

        st.markdown("##### Tempat & Tanggal Penandatanganan Surat")
        col_st1, col_st2 = st.columns(2)
        with col_st1:
            s_tempat = st.text_input("Tempat Surat", value="Jakarta", key="s_tempat")
        with col_st2:
            s_tgl_surat = st.date_input("Tanggal Dokumen", value=datetime.date.today(), key="s_tgl_surat")

        s_submitted = st.form_submit_button("🚀 Buat Dokumen Pendaftaran Sidang (1 Lembar)", use_container_width=True)

    if s_submitted:
        if not s_nama.strip():
            st.error("⚠️ Nama Lengkap wajib diisi!")
        elif not s_nim.strip():
            st.error("⚠️ NIM wajib diisi!")
        else:
            tpl_sidang = os.path.join(os.path.dirname(__file__), "template_sidang.docx")
            if not os.path.exists(tpl_sidang):
                st.error("❌ File template_sidang.docx tidak ditemukan di direktori aplikasi!")
            else:
                payload_s = {
                    "nama": s_nama.strip(),
                    "nim": s_nim.strip(),
                    "prodi": s_prodi.strip(),
                    "tempat": s_tempat.strip() or "Jakarta",
                    "tanggal": format_tanggal_indo(s_tgl_surat)
                }

                with st.spinner("Sedang memproses dokumen pendaftaran sidang..."):
                    docx_bytes_s = generate_sidang_document(tpl_sidang, payload_s)
                    pdf_bytes_s = convert_docx_to_pdf(docx_bytes_s)
                    clean_nim_s = "".join(c for c in s_nim if c.isalnum())

                st.success("✅ Formulir Pendaftaran Sidang berhasil dibuat tepat 1 lembar A4!")
                st.markdown("##### Pilih format unduhan:")
                col_sdl1, col_sdl2 = st.columns(2)
                with col_sdl1:
                    if pdf_bytes_s:
                        st.download_button(
                            label="📄 Unduh PDF (Pas 1 Lembar)",
                            data=pdf_bytes_s,
                            file_name=f"Pendaftaran_Sidang_Skripsi_{clean_nim_s}.pdf",
                            mime="application/pdf",
                            use_container_width=True
                        )
                    else:
                        st.warning("Konversi PDF tidak tersedia.")
                with col_sdl2:
                    st.download_button(
                        label="📝 Unduh Word (.docx)",
                        data=docx_bytes_s,
                        file_name=f"Pendaftaran_Sidang_Skripsi_{clean_nim_s}.docx",
                        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                        use_container_width=True
                    )
