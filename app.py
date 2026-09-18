import datetime
import hashlib
import hmac
import io
import os
import re
import shutil
import subprocess
import tempfile
import urllib.parse
from docx import Document
from docx.shared import Mm, Pt, RGBColor
from docx.oxml import parse_xml
from docx.oxml.ns import nsdecls
import pymupdf
from PIL import Image
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

def render_saweria_box(is_downloaded: bool = False):
    st.markdown("---")
    if is_downloaded:
        st.success("🎉 **Berkas berhasil diunduh!** Semoga urusan administrasi dan munaqasyahnya lancar berkah! 🎓✨")
    
    st.markdown(
        """
        <div style="background-color: #f8fafc; border: 1px solid #e2e8f0; border-radius: 12px; padding: 18px 20px; margin-top: 10px; margin-bottom: 12px;">
            <div style="display: flex; align-items: center; gap: 8px; margin-bottom: 6px;">
                <span style="font-size: 22px;">☕</span>
                <span style="font-size: 16px; font-weight: 700; color: #1e293b;">Selesai unduh berkas? Traktir kopi yuk!</span>
            </div>
            <p style="margin: 0 0 12px 0; color: #475569; font-size: 13.5px; line-height: 1.5;">
                Aplikasi ini dibuat gratis dan sukarela untuk mempermudah administrasi teman-teman mahasiswa SPI UIN Jakarta.<br>
                Jika merasa terbantu dan dokumenmu rapi tepat 1 lembar tanpa repot edit manual, kamu bisa traktir kopi pengembang secara sukarela via <b>Saweria</b> (bisa scan QRIS GoPay, DANA, OVO, ShopeePay, dan Bank).
            </p>
        </div>
        """,
        unsafe_allow_html=True
    )
    st.link_button(
        label="💛 Traktir Kopi Pengembang via Saweria (https://saweria.co/syarifhiday)",
        url="https://saweria.co/syarifhiday",
        use_container_width=True
    )

def get_user_field(key_names: list[str], default: str = "") -> str:
    """Mengambil nilai tersimpan dari form lain agar mahasiswa tidak perlu ketik ulang."""
    for k in key_names:
        val = st.session_state.get(k)
        if val and isinstance(val, str) and val.strip():
            return val.strip()
    return default

# ==========================================
# FUNGSI GENERATOR DOKUMEN (WORD & TEMPLATE)
# ==========================================
def generate_penyerahan_document(template_path: str, data: dict) -> bytes:
    doc = Document(template_path)
    sec = doc.sections[0]
    sec.top_margin = Mm(15)
    sec.bottom_margin = Mm(12)
    sec.left_margin = Mm(20)
    sec.right_margin = Mm(20)

    doc.paragraphs[1].runs[-1].text = f"\t: {data['nama']}"
    doc.paragraphs[2].runs[-1].text = f"\t: {data['nim']}"
    if data.get("prodi"):
        doc.paragraphs[3].runs[-1].text = data["prodi"]

    doc.paragraphs[4].runs[-1].text = f": {data['tanggal_sidang']}"
    doc.paragraphs[4].runs[-1].font.color.rgb = RGBColor(0, 0, 0)
    doc.paragraphs[5].runs[-1].text = f"\t: {data['judul']}"

    p7 = doc.paragraphs[7]._p
    p7.getparent().remove(p7)
    p6 = doc.paragraphs[6]
    p6.paragraph_format.space_before = Pt(0)
    p6.paragraph_format.space_after = Pt(0)
    p6.paragraph_format.line_spacing = 1.0

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

    p_tabs = doc.paragraphs[7]._p
    p_tabs.getparent().remove(p_tabs)

    p_jakarta = doc.paragraphs[7]
    tempat_tgl = f"{data['tempat']}, {data['tanggal_penyerahan']}"
    p_jakarta.runs[0].text = tempat_tgl
    p_jakarta.runs[0].font.color.rgb = RGBColor(0, 0, 0)
    for r in p_jakarta.runs[1:]:
        r.text = ""

    p_tabs2 = doc.paragraphs[8]._p
    p_tabs2.getparent().remove(p_tabs2)

    p_sig1 = doc.paragraphs[10]._p
    p_sig1.getparent().remove(p_sig1)
    p_sig2 = doc.paragraphs[10]._p
    p_sig2.getparent().remove(p_sig2)

    p_iwan = doc.paragraphs[10]
    p_iwan.paragraph_format.space_before = Pt(40)

    output_stream = io.BytesIO()
    doc.save(output_stream)
    return output_stream.getvalue()

def generate_ujian_document(template_path: str, data: dict) -> bytes:
    doc = Document(template_path)
    doc.paragraphs[1].runs[-1].text = f"\t: {data['nama']}"
    doc.paragraphs[2].runs[-1].text = f"\t: {data['nim']}"
    doc.paragraphs[3].runs[-1].text = f"\t: {data['prodi']}"
    doc.paragraphs[4].runs[-1].text = f"\t: {data['ttl']}"
    doc.paragraphs[5].runs[-1].text = f"\t: {data['asal_slta']}"
    doc.paragraphs[6].runs[-1].text = f"\t: {data['alamat']}"
    doc.paragraphs[7].runs[-1].text = f"\t: {data['ipk']}"
    doc.paragraphs[8].runs[-1].text = f"\t: {data['judul']}"
    doc.paragraphs[9].runs[-1].text = f": {data['kontak']}"
    doc.paragraphs[11].runs[0].text = f"Dengan ini mengajukan permohonan ujian skripsi pada semester {data['semester']} Tahun Akademik {data['tahun_akademik']}"
    doc.paragraphs[13].runs[0].text = f"{data['tempat']}, {data['tanggal']}"
    for r in doc.paragraphs[13].runs[1:]:
        r.text = ""
    doc.paragraphs[16].runs[0].text = data["nama"]
    doc.paragraphs[17].runs[0].text = f"NIM : {data['nim']}"

    for i in range(len(doc.paragraphs) - 1, -1, -1):
        txt = doc.paragraphs[i].text.strip()
        if txt.startswith("Catatan") or "Tanda Bintang" in txt:
            p_el = doc.paragraphs[i]._p
            p_el.getparent().remove(p_el)

    output_stream = io.BytesIO()
    doc.save(output_stream)
    return output_stream.getvalue()

def generate_sidang_document(template_path: str, data: dict) -> bytes:
    doc = Document(template_path)
    doc.paragraphs[0].runs[-1].text = f"\t: {data['nama']}"
    doc.paragraphs[1].runs[-1].text = f"\t: {data['nim']}"
    doc.paragraphs[2].runs[-1].text = f"\t: {data['prodi']}"
    doc.paragraphs[19].runs[0].text = f"\t{data['tempat']}, {data['tanggal']}"
    for r in doc.paragraphs[19].runs[1:]:
        r.text = ""
    doc.paragraphs[24].runs[-1].text = "Iwan Kurniawan, S.Pd., M.Si."

    output_stream = io.BytesIO()
    doc.save(output_stream)
    return output_stream.getvalue()

def generate_bimbingan_document(template_path: str, data: dict) -> bytes:
    doc = Document(template_path)
    replacements = {
        "{nama}": data["nama"],
        "{nim}": data["nim"],
        "{prodi}": data["prodi"],
        "{judul}": data["judul"],
        "{tempat}": data["tempat"],
        "{tanggal}": data["tanggal"],
        "{kaprodi_nama}": data["kaprodi_nama"],
        "{kaprodi_nip}": data["kaprodi_nip"],
        "{pembimbing_nama}": data["pembimbing_nama"],
        "{pembimbing_nip}": data["pembimbing_nip"]
    }
    for t in doc.tables:
        for row in t.rows:
            for cell in row.cells:
                for p in cell.paragraphs:
                    for r in p.runs:
                        for k, v in replacements.items():
                            if k in r.text:
                                r.text = r.text.replace(k, v)
    output_stream = io.BytesIO()
    doc.save(output_stream)
    return output_stream.getvalue()

def generate_pernyataan_document(template_path: str, data: dict) -> bytes:
    doc = Document(template_path)
    replacements = {
        "{nama}": data["nama"],
        "{nim}": data["nim"],
        "{prodi}": data["prodi"],
        "{tempat}": data["tempat"],
        "{tanggal}": data["tanggal"],
        "{materai_placeholder}": data.get("materai_placeholder", "")
    }
    for t in doc.tables:
        for row in t.rows:
            for cell in row.cells:
                for p in cell.paragraphs:
                    for r in p.runs:
                        for k, v in replacements.items():
                            if k in r.text:
                                r.text = r.text.replace(k, v)
    for p in doc.paragraphs:
        for r in p.runs:
            for k, v in replacements.items():
                if k in r.text:
                    r.text = r.text.replace(k, v)
    output_stream = io.BytesIO()
    doc.save(output_stream)
    return output_stream.getvalue()

def generate_publikasi_document(template_path: str, data: dict) -> bytes:
    doc = Document(template_path)
    replacements = {
        "{nama}": data["nama"],
        "{nim}": data["nim"],
        "{telepon}": data["telepon"],
        "{jurusan}": data["jurusan"],
        "{pembimbing}": data["pembimbing"],
        "{judul}": data["judul"],
        "{tempat}": data["tempat"],
        "{tanggal}": data["tanggal"]
    }
    for t in doc.tables:
        for row in t.rows:
            for cell in row.cells:
                for p in cell.paragraphs:
                    for r in p.runs:
                        for k, v in replacements.items():
                            if k in r.text:
                                r.text = r.text.replace(k, v)
    output_stream = io.BytesIO()
    doc.save(output_stream)
    return output_stream.getvalue()

# ==========================================
# WATERMARK PDF FUNCTIONS
# ==========================================
def parse_page_selection(total_pages: int, mode: str, skip_str: str = "", include_str: str = "") -> set[int]:
    """Mengembalikan set 0-indexed halaman yang AKAN di-watermark."""
    all_pages = set(range(total_pages))
    
    if mode == "Semua Halaman":
        return all_pages
    elif mode == "Semua Kecuali Cover (Halaman 1)":
        return {p for p in all_pages if p != 0}
    
    def parse_ranges(s: str) -> set[int]:
        res = set()
        if not s or not s.strip():
            return res
        parts = re.split(r"[,;]+", s.strip())
        for p in parts:
            p = p.strip()
            if not p:
                continue
            if "-" in p:
                sub = p.split("-")
                try:
                    start = int(sub[0].strip())
                    end = int(sub[1].strip())
                    for page_num in range(start, end + 1):
                        if 1 <= page_num <= total_pages:
                            res.add(page_num - 1)
                except ValueError:
                    continue
            else:
                try:
                    page_num = int(p)
                    if 1 <= page_num <= total_pages:
                        res.add(page_num - 1)
                except ValueError:
                    continue
        return res

    if include_str.strip():
        target = parse_ranges(include_str)
    else:
        target = all_pages

    skip = parse_ranges(skip_str)
    return target - skip


def prepare_watermark_image(raw_bytes: bytes, opacity: float, remove_white_bg: bool = False) -> bytes:
    """Mempersiapkan image watermark PNG dengan opacity dan pilihan hapus background putih."""
    img = Image.open(io.BytesIO(raw_bytes)).convert("RGBA")
    
    if remove_white_bg:
        datas = img.getdata()
        new_data = []
        for item in datas:
            if item[0] > 235 and item[1] > 235 and item[2] > 235:
                new_data.append((255, 255, 255, 0))
            else:
                new_data.append(item)
        img.putdata(new_data)
        
    r, g, b, a = img.split()
    a = a.point(lambda p: int(p * opacity))
    img.putalpha(a)
    
    out = io.BytesIO()
    img.save(out, format="PNG")
    return out.getvalue()


def apply_watermark_to_pdf(
    pdf_bytes: bytes,
    watermark_png_bytes: bytes,
    pages_to_watermark: set[int],
    scale: float = 0.45,
    sample_only: int | None = None
) -> tuple[bytes, list[bytes]]:
    """
    Menghasilkan PDF ber-watermark dan list bytes gambar PNG untuk sampel 5 halaman.
    """
    src_doc = pymupdf.open(stream=pdf_bytes, filetype="pdf")
    total_pages = len(src_doc)
    
    pil_img = Image.open(io.BytesIO(watermark_png_bytes))
    img_w, img_h = pil_img.size
    aspect = img_h / img_w
    
    sample_images = []
    
    if sample_only is not None:
        target_indices = list(range(min(sample_only, total_pages)))
        out_doc = pymupdf.open()
        for idx in target_indices:
            out_doc.insert_pdf(src_doc, from_page=idx, to_page=idx)
        proc_doc = out_doc
    else:
        target_indices = list(range(total_pages))
        proc_doc = src_doc

    for i, page in enumerate(proc_doc):
        orig_page_idx = i if sample_only is None else target_indices[i]
        
        if orig_page_idx in pages_to_watermark:
            rect = page.rect
            target_w = rect.width * scale
            target_h = target_w * aspect
            
            if target_h > rect.height * scale:
                target_h = rect.height * scale
                target_w = target_h / aspect
                
            x0 = (rect.width - target_w) / 2
            y0 = (rect.height - target_h) / 2
            x1 = x0 + target_w
            y1 = y0 + target_h
            
            img_rect = pymupdf.Rect(x0, y0, x1, y1)
            page.insert_image(img_rect, stream=watermark_png_bytes, overlay=False)
            
        if len(sample_images) < 5:
            pix = page.get_pixmap(dpi=110)
            sample_images.append(pix.tobytes("png"))

    out_bytes = proc_doc.tobytes()
    return out_bytes, sample_images


# Header Utama Aplikasi
st.title("🎓 Portal Layanan Berkas Skripsi")
st.subheader("Program Studi Sejarah dan Peradaban Islam (SPI)")
st.caption("Fakultas Adab dan Humaniora — UIN Syarif Hidayatullah Jakarta")

with st.sidebar:
    st.markdown("### 🎓 Layanan Berkas SPI")
    st.caption("Fakultas Adab dan Humaniora — UIN Syarif Hidayatullah Jakarta")
    st.markdown("---")
    st.markdown("### ☕ Dukung Pengembang")
    st.caption("Aplikasi ini dibuat sukarela untuk mempermudah administrasi skripsi mahasiswa SPI tepat 1 lembar A4.")
    st.link_button(
        label="💛 Donasi via Saweria",
        url="https://saweria.co/syarifhiday",
        use_container_width=True
    )

# 3 Tab Utama Berdasarkan Alur Mahasiswa (Sangat Responsif & Nyaman di HP)
tab1, tab2, tab3, tab4 = st.tabs([
    "📌 Panduan & Alur Berkas",
    "📝 Berkas Pendaftaran (Pra-Sidang)",
    "🎓 Berkas Kelulusan (Pasca-Sidang)",
    "💧 Watermark PDF Skripsi"
])

# ==========================================
# TAB 1: PANDUAN & CHECKLIST PERSYARATAN
# ==========================================
with tab1:
    st.info("💡 Informasi resmi berkas persyaratan ujian skripsi dan pengurusan BAP berdasarkan panduan Program Studi Sejarah dan Peradaban Islam (SPI) FAH UIN Syarif Hidayatullah Jakarta.")
    
    st.markdown("### 📋 1. Persyaratan Sidang Skripsi (12 Berkas Map Kuning)")
    st.caption("Semua berkas persyaratan ini dimasukkan ke dalam map kuning dan dikirimkan juga melalui email ke prodi.")
    
    col_chk1, col_chk2 = st.columns(2)
    with col_chk1:
        st.checkbox("1. Formulir Pendaftaran Sidang Skripsi *(tersedia di Tab 📝 Pra-Sidang)*")
        st.checkbox("2. Nilai IPK yang sudah dilegalisir")
        st.checkbox("3. Rekapitulasi Pembayaran Semester yang dilegalisir")
        st.checkbox("4. Lembar Pengesahan Skripsi (1 lembar)")
        st.checkbox("5. Fotokopi Ijazah SMA/SLTA (1 lembar)")
        st.checkbox("6. 1 Bundel Skripsi Lengkap (format Word / .docx)")
    with col_chk2:
        st.checkbox("7. Sertifikat Lulus TOAFL dan TOEFL")
        st.checkbox("8. Lulus Praktik Ibadah dan Qiro'ah (sesuai KRS Semester 2)")
        st.checkbox("9. Fotokopi Sertifikat Propesa / PBAK")
        st.checkbox("10. Surat Pernyataan Skripsi *(tersedia di Tab 📝 Pra-Sidang)*")
        st.checkbox("11. Surat Pernyataan Keaslian Berkas")
        st.checkbox("12. Jurnal / Lembar Catatan Bimbingan Dosen *(tersedia di Tab 📝 Pra-Sidang)*")

    st.markdown("---")

    st.markdown("### 🎓 2. Persyaratan BAP & Transkrip Nilai (Pasca-Sidang)")
    st.caption("Alur pengurusan Berita Acara Pemeriksaan (BAP) dan Transkrip Nilai setelah selesai ujian munaqasyah:")
    
    st.markdown("""
    1. **Menyerahkan Hasil Revisi Skripsi:**
       - Menyerahkan naskah revisi yang sudah sesuai dengan pedoman penulisan ke Dosen Penguji dan Dosen Pembimbing.
       - Dibuktikan dengan **Lembar Tanda Bukti Penyerahan Skripsi** *(tersedia di Tab 🎓 Pasca-Sidang)*.
    2. **Menyerahkan Hasil Revisi Skripsi dalam Bentuk PDF:**
       - Dokumen skripsi dalam bentuk PDF yang sudah disusun sesuai panduan penulisan (*dengan watermark resmi UIN*).
    3. **Submit Artikel Jurnal Ilmiah:**
       - Mengirimkan bukti pengiriman / submit skripsi yang diformat menjadi artikel jurnal ilmiah ke:
         - Jurnal **Socio Historica** (Jurnal Ilmiah Prodi SPI): `https://journal.uinjkt.ac.id/index.php/sh`
         - Atau ke jurnal ilmiah terakreditasi lainnya.
    4. **Surat Pernyataan Izin Publikasi Repository:**
       - Menandatangani Surat Pernyataan Izin Publikasi di Repository Perpustakaan UIN di atas materai 10.000 *(tersedia di Tab 🎓 Pasca-Sidang)*.
    5. **Pengiriman Seluruh Berkas Bukti:**
       - Seluruh bukti dikirimkan melalui email resmi prodi *(Template email siap salin & buka di Gmail tersedia di Tab 🎓 Pasca-Sidang)*.
    """)

    col_btn_jurnal, col_btn_template = st.columns(2)
    with col_btn_jurnal:
        st.link_button(
            "🔗 Buka Jurnal Socio Historica UIN Jakarta",
            "https://journal.uinjkt.ac.id/index.php/sh",
            use_container_width=True
        )
    with col_btn_template:
        st.info("✉️ Format teks email resmi permohonan BAP bisa langsung disalin pada Tab 🎓 Pasca-Sidang.")

    st.markdown("---")
    st.markdown("### 📬 Alamat Email Resmi Program Studi SPI")
    st.markdown("""
    Pengiriman berkas persyaratan sidang maupun permohonan BAP ditujukan ke alamat email resmi prodi berikut:
    - 📧 **`ski.fah@apps.uinjkt.ac.id`** (Akun Resmi Google Apps UIN)
    - 📧 **`spi.fah.uinjakarta@gmail.com`** (Akun Cadangan Prodi SPI)
    """)

    render_saweria_box()

# ==========================================
# TAB 2: BERKAS PENDAFTARAN (PRA-SIDANG)
# ==========================================
with tab2:
    pilihan_pra = st.radio(
        "Pilih Dokumen Pra-Sidang yang Ingin Dibuat:",
        [
            "📝 Formulir Permohonan Ujian Skripsi (TU)",
            "📋 Formulir Pendaftaran Sidang (Map Kuning)",
            "📖 Lembar Bimbingan Skripsi",
            "✍️ Lembar Pernyataan Skripsi"
        ],
        horizontal=True
    )
    st.markdown("---")

    # Nilai default otomatis dari input sebelumnya
    auto_nama = get_user_field(["u_nama", "s_nama", "b_nama", "per_nama", "p_nama", "pub_nama", "e_nama"])
    auto_nim = get_user_field(["u_nim", "s_nim", "b_nim", "per_nim", "p_nim", "pub_nim", "e_nim"])
    auto_prodi = get_user_field(["u_prodi", "s_prodi", "b_prodi", "per_prodi", "p_prodi", "pub_prodi", "e_prodi"], default="Sejarah dan Peradaban Islam")
    auto_judul = get_user_field(["u_judul", "b_judul", "p_judul", "pub_judul", "e_judul"])

    # ----------------------------------------------------
    # DOKUMEN 1 PRA: FORMULIR UJIAN SKRIPSI (TU)
    # ----------------------------------------------------
    if "Formulir Permohonan Ujian Skripsi" in pilihan_pra:
        st.info("💡 Formulir permohonan ujian skripsi (munaqasyah) untuk diserahkan ke bagian Tata Usaha (TU) Fakultas.")
        
        with st.form("form_ujian"):
            col_u1, col_u2 = st.columns(2)
            with col_u1:
                u_nama = st.text_input("Nama Lengkap", value=auto_nama, placeholder="Contoh: Syarif Hidayatullah", key="u_nama")
                u_ttl = st.text_input("Tempat/Tanggal Lahir", placeholder="Contoh: Tangerang, 12 Agustus 2002", key="u_ttl")
                u_alamat = st.text_area("Alamat Sekarang", placeholder="Contoh: Jl. Ciputat Raya No. 45, Tangerang Selatan", height=70, key="u_alamat")
            with col_u2:
                u_nim = st.text_input("NIM", value=auto_nim, placeholder="Contoh: 11200210000088", key="u_nim")
                u_prodi = st.text_input("Program Studi", value=auto_prodi, key="u_prodi")
                u_asal_slta = st.text_input("Asal SLTA / SMA", placeholder="Contoh: SMAN 1 Tangerang", key="u_slta")

            col_u3, col_u4 = st.columns(2)
            with col_u3:
                u_ipk = st.text_input("IPK Sementara", placeholder="Contoh: 3.75", key="u_ipk")
            with col_u4:
                u_kontak = st.text_input("No. Telpon / Email", placeholder="Contoh: 08123456789 / nama@gmail.com", key="u_kontak")

            u_judul = st.text_area("Judul Skripsi", value=auto_judul, placeholder="Tuliskan judul lengkap skripsi...", height=80, key="u_judul")

            col_u5, col_u6 = st.columns(2)
            with col_u5:
                u_semester = st.selectbox("Semester Pengajuan Ujian", ["Ganjil", "Genap"], key="u_sem")
            with col_u6:
                current_year = datetime.date.today().year
                default_ta = f"{current_year}/{current_year + 1}"
                u_ta = st.text_input("Tahun Akademik", value=default_ta, placeholder="Contoh: 2025/2026", key="u_ta")

            st.markdown("##### Tempat & Tanggal Dokumen")
            col_ut1, col_ut2 = st.columns(2)
            with col_ut1:
                u_tempat = st.text_input("Tempat Surat", value="Jakarta", key="u_tempat")
            with col_ut2:
                u_tanggal = st.date_input("Tanggal Surat", value=datetime.date.today(), key="u_tgl")

            u_submitted = st.form_submit_button("🚀 Buat Formulir Ujian Skripsi (1 Lembar)", use_container_width=True)

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
                    payload_u = {
                        "nama": u_nama.strip(),
                        "nim": u_nim.strip(),
                        "prodi": u_prodi.strip(),
                        "ttl": u_ttl.strip() or "-",
                        "asal_slta": u_asal_slta.strip() or "-",
                        "alamat": u_alamat.strip() or "-",
                        "ipk": u_ipk.strip() or "-",
                        "judul": u_judul.strip(),
                        "kontak": u_kontak.strip() or "-",
                        "semester": u_semester,
                        "tahun_akademik": u_ta.strip(),
                        "tempat": u_tempat.strip() or "Jakarta",
                        "tanggal": format_tanggal_indo(u_tanggal)
                    }

                    with st.spinner("Sedang memproses dokumen permohonan ujian..."):
                        docx_bytes_u = generate_ujian_document(tpl_ujian, payload_u)
                        pdf_bytes_u = convert_docx_to_pdf(docx_bytes_u)
                        clean_nim_u = "".join(c for c in u_nim if c.isalnum())
                        st.session_state["u_doc"] = {
                            "docx": docx_bytes_u,
                            "pdf": pdf_bytes_u,
                            "nim": clean_nim_u
                        }

        if "u_doc" in st.session_state:
            u_data = st.session_state["u_doc"]
            st.success("✅ Formulir Pendaftaran Ujian Skripsi berhasil dibuat tepat 1 lembar A4!")
            st.markdown("##### Pilih format unduhan:")
            col_udl1, col_udl2 = st.columns(2)
            with col_udl1:
                if u_data["pdf"]:
                    btn_u_pdf = st.download_button(
                        label="📄 Unduh PDF (Pas 1 Lembar)",
                        data=u_data["pdf"],
                        file_name=f"Pendaftaran_Ujian_Skripsi_{u_data['nim']}.pdf",
                        mime="application/pdf",
                        use_container_width=True,
                        key="dl_pdf_u"
                    )
                    if btn_u_pdf:
                        st.session_state["u_downloaded"] = True
                        st.toast("🎉 Formulir Ujian PDF berhasil diunduh! 🎓✨")
                        st.balloons()
                else:
                    st.warning("Konversi PDF tidak tersedia.")
            with col_udl2:
                btn_u_docx = st.download_button(
                    label="📝 Unduh Word (.docx)",
                    data=u_data["docx"],
                    file_name=f"Pendaftaran_Ujian_Skripsi_{u_data['nim']}.docx",
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                    use_container_width=True,
                    key="dl_docx_u"
                )
                if btn_u_docx:
                    st.session_state["u_downloaded"] = True
                    st.toast("🎉 Formulir Ujian Word berhasil diunduh! 🎓✨")
                    st.balloons()

            render_saweria_box(st.session_state.get("u_downloaded", False))

    # ----------------------------------------------------
    # DOKUMEN 2 PRA: FORMULIR PENDAFTARAN SIDANG (MAP KUNING)
    # ----------------------------------------------------
    elif "Formulir Pendaftaran Sidang" in pilihan_pra:
        st.info("💡 Formulir pendaftaran sidang skripsi resmi ber-kop Fakultas Adab dan Humaniora beserta checklist 11 berkas map kuning.")
        
        with st.form("form_sidang"):
            col_s1, col_s2 = st.columns(2)
            with col_s1:
                s_nama = st.text_input("Nama Lengkap Mahasiswa", value=auto_nama, placeholder="Contoh: Syarif Hidayatullah", key="s_nama")
                s_nim = st.text_input("NIM", value=auto_nim, placeholder="Contoh: 11200210000088", key="s_nim")
            with col_s2:
                s_prodi = st.text_input("Program Studi", value=auto_prodi, key="s_prodi")
                s_tempat = st.text_input("Tempat Surat", value="Jakarta", key="s_tempat")

            s_tgl_surat = st.date_input("Tanggal Surat", value=datetime.date.today(), key="s_tgl")

            st.caption("ℹ️ Dokumen ini akan menampilkan kop surat resmi Fakultas Adab dan Humaniora, 11 daftar persyaratan ujian skripsi, dan tanda tangan resmi Kabag TU.")
            s_submitted = st.form_submit_button("🚀 Buat Formulir Pendaftaran Sidang (1 Lembar)", use_container_width=True)

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
                        st.session_state["s_doc"] = {
                            "docx": docx_bytes_s,
                            "pdf": pdf_bytes_s,
                            "nim": clean_nim_s
                        }

        if "s_doc" in st.session_state:
            s_data = st.session_state["s_doc"]
            st.success("✅ Formulir Pendaftaran Sidang berhasil dibuat tepat 1 lembar A4!")
            st.markdown("##### Pilih format unduhan:")
            col_sdl1, col_sdl2 = st.columns(2)
            with col_sdl1:
                if s_data["pdf"]:
                    btn_s_pdf = st.download_button(
                        label="📄 Unduh PDF (Pas 1 Lembar)",
                        data=s_data["pdf"],
                        file_name=f"Pendaftaran_Sidang_Skripsi_{s_data['nim']}.pdf",
                        mime="application/pdf",
                        use_container_width=True,
                        key="dl_pdf_s"
                    )
                    if btn_s_pdf:
                        st.session_state["s_downloaded"] = True
                        st.toast("🎉 Formulir Sidang PDF berhasil diunduh! 🎓✨")
                        st.balloons()
                else:
                    st.warning("Konversi PDF tidak tersedia.")
            with col_sdl2:
                btn_s_docx = st.download_button(
                    label="📝 Unduh Word (.docx)",
                    data=s_data["docx"],
                    file_name=f"Pendaftaran_Sidang_Skripsi_{s_data['nim']}.docx",
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                    use_container_width=True,
                    key="dl_docx_s"
                )
                if btn_s_docx:
                    st.session_state["s_downloaded"] = True
                    st.toast("🎉 Formulir Sidang Word berhasil diunduh! 🎓✨")
                    st.balloons()

            render_saweria_box(st.session_state.get("s_downloaded", False))

    # ----------------------------------------------------
    # DOKUMEN 3 PRA: LEMBAR BIMBINGAN SKRIPSI
    # ----------------------------------------------------
    elif "Lembar Bimbingan Skripsi" in pilihan_pra:
        st.info("💡 Lembar Catatan Bimbingan Skripsi resmi prodi dengan tabel log bimbingan 8 baris (minimal 6 kali bimbingan) lengkap dengan tanda tangan Ketua Program Studi dan Dosen Pembimbing.")

        with st.form("form_bimbingan"):
            col_b1, col_b2 = st.columns(2)
            with col_b1:
                b_nama = st.text_input("Nama Lengkap Mahasiswa", value=auto_nama, placeholder="Contoh: Syarif Hidayatullah", key="b_nama")
                b_nim = st.text_input("NIM", value=auto_nim, placeholder="Contoh: 11210220000073", key="b_nim")
            with col_b2:
                b_prodi = st.text_input("Program Studi", value=auto_prodi, key="b_prodi")
                b_tempat = st.text_input("Tempat Surat", value="Jakarta", key="b_tempat")

            b_tgl = st.date_input("Tanggal Lembar Bimbingan", value=datetime.date.today(), key="b_tgl")
            b_judul = st.text_area("Judul Skripsi", value=auto_judul, placeholder="Tuliskan judul lengkap skripsi...", height=80, key="b_judul")

            st.markdown("##### Dosen Pembimbing Skripsi")
            st.caption("Ketik angka nomornya saja. Sistem otomatis mendeteksi: **18 digit = NIP**, **10 digit = NIDN**.")
            col_bp1, col_bp2 = st.columns(2)
            with col_bp1:
                b_pembimbing_nama = st.text_input("Nama & Gelar Dosen Pembimbing", placeholder="Contoh: Dr. Halimatus Sa'diyah, M.A.", key="b_pnama")
            with col_bp2:
                b_pembimbing_nip = st.text_input("Nomor NIP / NIDN Dosen Pembimbing", placeholder="Contoh: 198203152009012011", key="b_pnip")

            with st.expander("⚙️ Data Ketua Program Studi (Default: Kaprodi SPI)"):
                col_bk1, col_bk2 = st.columns(2)
                with col_bk1:
                    b_kaprodi_nama = st.text_input("Nama Kaprodi", value="Dr. Zakiya Darojat, M.A.", key="b_knama")
                with col_bk2:
                    b_kaprodi_nip = st.text_input("NIP/NIDN Kaprodi", value="197405302005012006", key="b_knip")

            b_submitted = st.form_submit_button("🚀 Buat Lembar Bimbingan Skripsi (1 Lembar)", use_container_width=True)

        if b_submitted:
            if not b_nama.strip():
                st.error("⚠️ Nama Lengkap wajib diisi!")
            elif not b_nim.strip():
                st.error("⚠️ NIM wajib diisi!")
            elif not b_judul.strip():
                st.error("⚠️ Judul Skripsi wajib diisi!")
            else:
                tpl_bimbingan = os.path.join(os.path.dirname(__file__), "template_bimbingan.docx")
                if not os.path.exists(tpl_bimbingan):
                    st.error("❌ File template_bimbingan.docx tidak ditemukan di direktori aplikasi!")
                else:
                    payload_b = {
                        "nama": b_nama.strip(),
                        "nim": b_nim.strip(),
                        "prodi": b_prodi.strip() or "Sejarah dan Peradaban Islam",
                        "judul": b_judul.strip(),
                        "tempat": b_tempat.strip() or "Jakarta",
                        "tanggal": format_tanggal_indo(b_tgl),
                        "kaprodi_nama": b_kaprodi_nama.strip() or "Dr. Zakiya Darojat, M.A.",
                        "kaprodi_nip": format_nip_nidn(b_kaprodi_nip.strip()) if b_kaprodi_nip.strip() else "NIP. 197405302005012006",
                        "pembimbing_nama": b_pembimbing_nama.strip() or "Nama Dosen Pembimbing",
                        "pembimbing_nip": format_nip_nidn(b_pembimbing_nip.strip()) if b_pembimbing_nip.strip() else "NIP/NIDN. "
                    }

                    with st.spinner("Sedang memproses Lembar Bimbingan Skripsi..."):
                        docx_bytes_b = generate_bimbingan_document(tpl_bimbingan, payload_b)
                        pdf_bytes_b = convert_docx_to_pdf(docx_bytes_b)
                        clean_nim_b = "".join(c for c in b_nim if c.isalnum())
                        st.session_state["b_doc"] = {
                            "docx": docx_bytes_b,
                            "pdf": pdf_bytes_b,
                            "nim": clean_nim_b
                        }

        if "b_doc" in st.session_state:
            b_data = st.session_state["b_doc"]
            st.success("✅ Lembar Bimbingan Skripsi berhasil dibuat tepat 1 lembar A4!")
            st.markdown("##### Pilih format unduhan:")
            col_bdl1, col_bdl2 = st.columns(2)
            with col_bdl1:
                if b_data["pdf"]:
                    btn_b_pdf = st.download_button(
                        label="📄 Unduh PDF (Pas 1 Lembar)",
                        data=b_data["pdf"],
                        file_name=f"Lembar_Bimbingan_Skripsi_{b_data['nim']}.pdf",
                        mime="application/pdf",
                        use_container_width=True,
                        key="dl_pdf_b"
                    )
                    if btn_b_pdf:
                        st.session_state["b_downloaded"] = True
                        st.toast("🎉 Lembar Bimbingan PDF berhasil diunduh! 🎓✨")
                        st.balloons()
                else:
                    st.warning("Konversi PDF tidak tersedia.")
            with col_bdl2:
                btn_b_docx = st.download_button(
                    label="📝 Unduh Word (.docx)",
                    data=b_data["docx"],
                    file_name=f"Lembar_Bimbingan_Skripsi_{b_data['nim']}.docx",
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                    use_container_width=True,
                    key="dl_docx_b"
                )
                if btn_b_docx:
                    st.session_state["b_downloaded"] = True
                    st.toast("🎉 Lembar Bimbingan Word berhasil diunduh! 🎓✨")
                    st.balloons()

            render_saweria_box(st.session_state.get("b_downloaded", False))

    # ----------------------------------------------------
    # DOKUMEN 4 PRA: LEMBAR PERNYATAAN SKRIPSI (BEBAS PLAGIASI)
    # ----------------------------------------------------
    elif "Lembar Pernyataan Skripsi" in pilihan_pra:
        st.info("💡 Lembar Pernyataan Skripsi (Bebas Plagiasi & Orisinalitas Penelitian) sesuai format resmi akademik UIN Syarif Hidayatullah Jakarta tepat 1 lembar A4.")

        with st.form("form_pernyataan"):
            col_per1, col_per2 = st.columns(2)
            with col_per1:
                per_nama = st.text_input("Nama Lengkap Mahasiswa", value=auto_nama, placeholder="Contoh: Syarif Hidayatullah", key="per_nama")
                per_nim = st.text_input("NIM", value=auto_nim, placeholder="Contoh: 11210220000073", key="per_nim")
            with col_per2:
                per_prodi = st.text_input("Program Studi", value=auto_prodi, key="per_prodi")
                per_tempat = st.text_input("Tempat Surat", value="Jakarta", key="per_tempat")

            per_tgl = st.date_input("Tanggal Surat Pernyataan", value=datetime.date.today(), key="per_tgl")
            per_materai = st.checkbox("Sediakan panduan teks Materai 10.000 (Opsional)", value=False, help="Centang bila ingin mencantumkan panduan posisi tempel materai 10.000 di atas tanda tangan", key="per_mat")

            per_submitted = st.form_submit_button("🚀 Buat Lembar Pernyataan Skripsi (1 Lembar)", use_container_width=True)

        if per_submitted:
            if not per_nama.strip():
                st.error("⚠️ Nama Lengkap wajib diisi!")
            elif not per_nim.strip():
                st.error("⚠️ NIM wajib diisi!")
            else:
                tpl_pernyataan = os.path.join(os.path.dirname(__file__), "template_pernyataan.docx")
                if not os.path.exists(tpl_pernyataan):
                    st.error("❌ File template_pernyataan.docx tidak ditemukan di direktori aplikasi!")
                else:
                    mat_text = "Materai 10.000" if per_materai else ""
                    payload_per = {
                        "nama": per_nama.strip(),
                        "nim": per_nim.strip(),
                        "prodi": per_prodi.strip() or "Sejarah Peradaban Islam",
                        "tempat": per_tempat.strip() or "Jakarta",
                        "tanggal": format_tanggal_indo(per_tgl),
                        "materai_placeholder": mat_text
                    }

                    with st.spinner("Sedang memproses Lembar Pernyataan Skripsi..."):
                        docx_bytes_per = generate_pernyataan_document(tpl_pernyataan, payload_per)
                        pdf_bytes_per = convert_docx_to_pdf(docx_bytes_per)
                        clean_nim_per = "".join(c for c in per_nim if c.isalnum())
                        st.session_state["per_doc"] = {
                            "docx": docx_bytes_per,
                            "pdf": pdf_bytes_per,
                            "nim": clean_nim_per
                        }

        if "per_doc" in st.session_state:
            per_data = st.session_state["per_doc"]
            st.success("✅ Lembar Pernyataan Skripsi berhasil dibuat tepat 1 lembar A4!")
            st.markdown("##### Pilih format unduhan:")
            col_perdl1, col_perdl2 = st.columns(2)
            with col_perdl1:
                if per_data["pdf"]:
                    btn_per_pdf = st.download_button(
                        label="📄 Unduh PDF (Pas 1 Lembar)",
                        data=per_data["pdf"],
                        file_name=f"Lembar_Pernyataan_Skripsi_{per_data['nim']}.pdf",
                        mime="application/pdf",
                        use_container_width=True,
                        key="dl_pdf_per"
                    )
                    if btn_per_pdf:
                        st.session_state["per_downloaded"] = True
                        st.toast("🎉 Lembar Pernyataan PDF berhasil diunduh! 🎓✨")
                        st.balloons()
                else:
                    st.warning("Konversi PDF tidak tersedia.")
            with col_perdl2:
                btn_per_docx = st.download_button(
                    label="📝 Unduh Word (.docx)",
                    data=per_data["docx"],
                    file_name=f"Lembar_Pernyataan_Skripsi_{per_data['nim']}.docx",
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                    use_container_width=True,
                    key="dl_docx_per"
                )
                if btn_per_docx:
                    st.session_state["per_downloaded"] = True
                    st.toast("🎉 Lembar Pernyataan Word berhasil diunduh! 🎓✨")
                    st.balloons()

            render_saweria_box(st.session_state.get("per_downloaded", False))

# ==========================================
# TAB 3: BERKAS KELULUSAN (PASCA-SIDANG)
# ==========================================
with tab3:
    pilihan_pasca = st.radio(
        "Pilih Layanan Pasca-Sidang / Kelulusan:",
        [
            "📄 Tanda Bukti Penyerahan Skripsi",
            "🏛️ Surat Izin Publikasi Repository",
            "✉️ Template Email Siap Salin"
        ],
        horizontal=True
    )
    st.markdown("---")

    auto_nama_pasca = get_user_field(["p_nama", "pub_nama", "e_nama", "s_nama", "u_nama", "b_nama", "per_nama"])
    auto_nim_pasca = get_user_field(["p_nim", "pub_nim", "e_nim", "s_nim", "u_nim", "b_nim", "per_nim"])
    auto_prodi_pasca = get_user_field(["p_prodi", "pub_jurusan", "e_prodi", "s_prodi", "u_prodi", "b_prodi", "per_prodi"], default="Sejarah dan Peradaban Islam")
    auto_judul_pasca = get_user_field(["p_judul", "pub_judul", "e_judul", "u_judul", "b_judul"])
    auto_pembimbing_pasca = get_user_field(["p_pnama", "pub_pembimbing", "b_pnama", "e_pembimbing"])

    # ----------------------------------------------------
    # DOKUMEN 1 PASCA: TANDA BUKTI PENYERAHAN SKRIPSI
    # ----------------------------------------------------
    if "Tanda Bukti Penyerahan Skripsi" in pilihan_pasca:
        st.info("💡 Formulir tanda bukti penyerahan skripsi pasca munaqasyah ke pihak Fakultas, Prodi, dan Perpustakaan.")
        
        with st.form("form_penyerahan"):
            col1, col2 = st.columns(2)
            with col1:
                p_nama = st.text_input("Nama Lengkap Mahasiswa", value=auto_nama_pasca, placeholder="Contoh: Syarif Hidayatullah", key="p_nama")
            with col2:
                p_nim = st.text_input("NIM", value=auto_nim_pasca, placeholder="Contoh: 11200210000088", key="p_nim")

            p_prodi = st.text_input("Program Studi", value=auto_prodi_pasca, key="p_prodi")
            p_tgl_sidang = st.date_input("Tanggal Sidang Munaqasyah", value=datetime.date.today(), key="p_sidang")
            p_judul = st.text_area("Judul Skripsi", value=auto_judul_pasca, placeholder="Tuliskan judul skripsi naskah final...", height=90, key="p_judul")

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
                    p_pembimbing_nama = st.text_input("Nama & Gelar Pembimbing", value=auto_pembimbing_pasca, placeholder="Nama Pembimbing", key="p_pnama")
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

                    with st.spinner("Sedang memproses dokumen tanda penyerahan skripsi..."):
                        docx_bytes_p = generate_penyerahan_document(tpl_penyerahan, payload_p)
                        pdf_bytes_p = convert_docx_to_pdf(docx_bytes_p)
                        clean_nim_p = "".join(c for c in p_nim if c.isalnum())
                        st.session_state["p_doc"] = {
                            "docx": docx_bytes_p,
                            "pdf": pdf_bytes_p,
                            "nim": clean_nim_p
                        }

        if "p_doc" in st.session_state:
            p_data = st.session_state["p_doc"]
            st.success("✅ Dokumen berhasil dibuat tepat 1 lembar A4!")
            st.markdown("##### Pilih format unduhan:")
            col_p1, col_p2 = st.columns(2)
            with col_p1:
                if p_data["pdf"]:
                    btn_p_pdf = st.download_button(
                        label="📄 Unduh PDF (Pas 1 Lembar)",
                        data=p_data["pdf"],
                        file_name=f"Tanda_Bukti_Penyerahan_Skripsi_{p_data['nim']}.pdf",
                        mime="application/pdf",
                        use_container_width=True,
                        key="dl_pdf_p"
                    )
                    if btn_p_pdf:
                        st.session_state["p_downloaded"] = True
                        st.toast("🎉 Berkas PDF berhasil diunduh! 🎓✨")
                        st.balloons()
                else:
                    st.warning("Konversi PDF tidak tersedia.")
            with col_p2:
                btn_p_docx = st.download_button(
                    label="📝 Unduh Word (.docx)",
                    data=p_data["docx"],
                    file_name=f"Tanda_Bukti_Penyerahan_Skripsi_{p_data['nim']}.docx",
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                    use_container_width=True,
                    key="dl_docx_p"
                )
                if btn_p_docx:
                    st.session_state["p_downloaded"] = True
                    st.toast("🎉 Berkas Word berhasil diunduh! 🎓✨")
                    st.balloons()

            render_saweria_box(st.session_state.get("p_downloaded", False))

    # ----------------------------------------------------
    # DOKUMEN 2 PASCA: SURAT PERNYATAAN IZIN PUBLIKASI
    # ----------------------------------------------------
    elif "Surat Izin Publikasi Repository" in pilihan_pasca:
        st.info("💡 Surat Pernyataan Izin Publikasi Karya Tulis Ilmiah (Skripsi) di Repository Perpustakaan UIN Syarif Hidayatullah Jakarta (ditandatangani di atas materai 10.000).")

        with st.form("form_publikasi"):
            col_pub1, col_pub2 = st.columns(2)
            with col_pub1:
                pub_nama = st.text_input("Nama Lengkap", value=auto_nama_pasca, placeholder="Contoh: Syarif Hidayatullah", key="pub_nama")
                pub_telepon = st.text_input("No. Telepon / HP", placeholder="Contoh: 081234567890", key="pub_telp")
            with col_pub2:
                pub_nim = st.text_input("NIM", value=auto_nim_pasca, placeholder="Contoh: 11200210000088", key="pub_nim")
                pub_jurusan = st.text_input("Jurusan / Program Studi", value=auto_prodi_pasca, key="pub_jurusan")

            pub_pembimbing = st.text_input("Dosen Pembimbing (Lengkap dengan Gelar)", value=auto_pembimbing_pasca, placeholder="Contoh: Dr. Nama Pembimbing, M.Hum.", key="pub_pembimbing")
            pub_judul = st.text_area("Judul Skripsi/Tesis (Latin)", value=auto_judul_pasca, placeholder="Tuliskan judul lengkap skripsi...", height=80, key="pub_judul")

            st.markdown("##### Tempat & Tanggal Penandatanganan")
            col_pt1, col_pt2 = st.columns(2)
            with col_pt1:
                pub_tempat = st.text_input("Tempat Surat", value="Jakarta", key="pub_tempat")
            with col_pt2:
                pub_tgl = st.date_input("Tanggal Surat", value=datetime.date.today(), key="pub_tgl")

            pub_submitted = st.form_submit_button("🚀 Buat Surat Izin Publikasi (1 Lembar)", use_container_width=True)

        if pub_submitted:
            if not pub_nama.strip():
                st.error("⚠️ Nama Lengkap wajib diisi!")
            elif not pub_nim.strip():
                st.error("⚠️ NIM wajib diisi!")
            elif not pub_judul.strip():
                st.error("⚠️ Judul Skripsi wajib diisi!")
            else:
                tpl_publikasi = os.path.join(os.path.dirname(__file__), "template_publikasi.docx")
                if not os.path.exists(tpl_publikasi):
                    st.error("❌ File template_publikasi.docx tidak ditemukan di direktori aplikasi!")
                else:
                    payload_pub = {
                        "nama": pub_nama.strip(),
                        "nim": pub_nim.strip(),
                        "telepon": pub_telepon.strip() or "-",
                        "jurusan": pub_jurusan.strip() or "Sejarah dan Peradaban Islam",
                        "pembimbing": pub_pembimbing.strip() or "-",
                        "judul": pub_judul.strip(),
                        "tempat": pub_tempat.strip() or "Jakarta",
                        "tanggal": format_tanggal_indo(pub_tgl)
                    }

                    with st.spinner("Sedang memproses Surat Pernyataan Izin Publikasi..."):
                        docx_bytes_pub = generate_publikasi_document(tpl_publikasi, payload_pub)
                        pdf_bytes_pub = convert_docx_to_pdf(docx_bytes_pub)
                        clean_nim_pub = "".join(c for c in pub_nim if c.isalnum())
                        st.session_state["pub_doc"] = {
                            "docx": docx_bytes_pub,
                            "pdf": pdf_bytes_pub,
                            "nim": clean_nim_pub
                        }

        if "pub_doc" in st.session_state:
            pub_data = st.session_state["pub_doc"]
            st.success("✅ Surat Pernyataan Izin Publikasi berhasil dibuat tepat 1 lembar A4!")
            st.markdown("##### Pilih format unduhan:")
            col_pdl1, col_pdl2 = st.columns(2)
            with col_pdl1:
                if pub_data["pdf"]:
                    btn_pub_pdf = st.download_button(
                        label="📄 Unduh PDF (Pas 1 Lembar)",
                        data=pub_data["pdf"],
                        file_name=f"Surat_Izin_Publikasi_{pub_data['nim']}.pdf",
                        mime="application/pdf",
                        use_container_width=True,
                        key="dl_pdf_pub"
                    )
                    if btn_pub_pdf:
                        st.session_state["pub_downloaded"] = True
                        st.toast("🎉 Surat Izin Publikasi PDF berhasil diunduh! 🎓✨")
                        st.balloons()
                else:
                    st.warning("Konversi PDF tidak tersedia.")
            with col_pdl2:
                btn_pub_docx = st.download_button(
                    label="📝 Unduh Word (.docx)",
                    data=pub_data["docx"],
                    file_name=f"Surat_Izin_Publikasi_{pub_data['nim']}.docx",
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                    use_container_width=True,
                    key="dl_docx_pub"
                )
                if btn_pub_docx:
                    st.session_state["pub_downloaded"] = True
                    st.toast("🎉 Surat Izin Publikasi Word berhasil diunduh! 🎓✨")
                    st.balloons()

            render_saweria_box(st.session_state.get("pub_downloaded", False))

    # ----------------------------------------------------
    # DOKUMEN 3 PASCA: TEMPLATE EMAIL SIAP SALIN
    # ----------------------------------------------------
    elif "Template Email Siap Salin" in pilihan_pasca:
        st.info("💡 Generator format email resmi ke Program Studi SPI FAH UIN Jakarta sesuai infografis resmi. Tinggal lengkapi identitas, salin dalam 1 klik, atau buka langsung di Gmail / aplikasi email!")

        pilihan_skenario = st.radio(
            "Pilih Jenis Permohonan Email:",
            [
                "1️⃣ Permohonan Pendaftaran Ujian Proposal / Skripsi (Pra-Sidang)",
                "2️⃣ Permohonan Berita Acara Sidang Skripsi (BAP & Transkrip Nilai)"
            ],
            horizontal=True
        )

        st.markdown("##### Lengkapi Data Pemohon:")
        col_e1, col_e2 = st.columns(2)
        with col_e1:
            e_nama = st.text_input("Nama Lengkap", value=auto_nama_pasca, placeholder="Contoh: Syarif Hidayatullah", key="e_nama")
            e_nim = st.text_input("NIM", value=auto_nim_pasca, placeholder="Contoh: 11200210000088", key="e_nim")
        with col_e2:
            e_prodi = st.text_input("Program Studi", value=auto_prodi_pasca, key="e_prodi")
            e_telepon = st.text_input("Nomor Telepon / WhatsApp", placeholder="Contoh: 081234567890", key="e_telp")

        e_judul = st.text_area("Judul Proposal / Skripsi", value=auto_judul_pasca, placeholder="Tuliskan judul proposal / skripsi...", height=70, key="e_judul")

        if "1️⃣" in pilihan_skenario:  # Permohonan Pendaftaran Ujian Proposal / Skripsi
            e_pembimbing = st.text_input("Nama Dosen Pembimbing (Lengkap dengan Gelar)", value=auto_pembimbing_pasca, placeholder="Contoh: Dr. Nama Pembimbing, M.Hum.", key="e_pembimbing")
            
            subject_text = f"Permohonan Pendaftaran Ujian Proposal/Skripsi - {e_nama or '[Nama]'} - {e_nim or '[NIM]'}"
            pembimbing_str = e_pembimbing.strip() if e_pembimbing.strip() else "[Nama Dosen Pembimbing]"
            
            body_text = f"""Assalamualaikum wr.wb,

Dengan hormat,
Saya yang bertanda tangan di bawah ini:
Nama: {e_nama.strip() or '[Nama Lengkap]'}
NIM: {e_nim.strip() or '[Nomor Induk Mahasiswa]'}
Program Studi: {e_prodi.strip() or 'Sejarah dan Peradaban Islam'}
Nomor Telepon: {e_telepon.strip() or '[Nomor Telepon]'}

Dengan ini, mengajukan permohonan untuk mendaftar ujian proposal/skripsi. Proposal Skripsi saya berjudul “{e_judul.strip() or '[Judul Proposal/Skripsi]'}” telah selesai dan telah mendapatkan persetujuan dari {pembimbing_str} selaku dosen pembimbing.

Bersama email ini turut saya melampirkan dokumen-dokumen yang menjadi persyaratan ujian.

Atas perhatiannya kami ucapkan terima kasih.

Wassalamualaikum wr.wb."""

        else:  # Permohonan Berita Acara Sidang Skripsi (BAP)
            subject_text = f"Permohonan Berita Acara Sidang Skripsi - {e_nama or '[Nama]'} - {e_nim or '[NIM]'}"
            
            body_text = f"""Assalamualaikum wr.wb,

Dengan hormat,
Saya yang bertanda tangan di bawah ini:
Nama: {e_nama.strip() or '[Nama Lengkap]'}
NIM: {e_nim.strip() or '[Nomor Induk Mahasiswa]'}
Program Studi: {e_prodi.strip() or 'Sejarah dan Peradaban Islam'}
Nomor Telepon: {e_telepon.strip() or '[Nomor Telepon]'}

Dengan ini, mengajukan permohonan Berita Acara Sidang Skripsi. Bersama email ini turut saya melampirkan dokumen-dokumen yang menjadi persyaratan yaitu :
1. Dokumen Skripsi dalam bentuk PDF yang sudah disusun sesuai panduan penulisan (dengan watermark)
2. Lembar Penyerahan Skripsi
3. Bukti Submit artikel Jurnal

Atas perhatiannya kami ucapkan terima kasih.

Wassalamualaikum wr.wb."""

        st.markdown("---")
        st.markdown("### 📤 Format Email Siap Kirim")
        
        email_tujuan = st.selectbox(
            "Pilih Alamat Email Tujuan Prodi:",
            [
                "ski.fah@apps.uinjkt.ac.id (Email Resmi Apps UIN)",
                "spi.fah.uinjakarta@gmail.com (Email Cadangan Prodi)",
                "ski.fah@apps.uinjkt.ac.id, spi.fah.uinjakarta@gmail.com (Kirim ke Keduanya)"
            ],
            key="e_target"
        )
        clean_target = email_tujuan.split()[0] if "," not in email_tujuan else "ski.fah@apps.uinjkt.ac.id,spi.fah.uinjakarta@gmail.com"

        st.markdown("**1. Subjek Email:**")
        st.code(subject_text, language="text")

        st.markdown("**2. Isi Pesan Email (Body):** *(Klik tombol copy di pojok kanan atas kotak)*")
        st.code(body_text, language="text")

        # Tombol Buka di Gmail / Mail Client
        mailto_url = f"mailto:{clean_target}?subject={urllib.parse.quote(subject_text)}&body={urllib.parse.quote(body_text)}"
        gmail_url = f"https://mail.google.com/mail/?view=cm&fs=1&to={clean_target}&su={urllib.parse.quote(subject_text)}&body={urllib.parse.quote(body_text)}"

        col_mail1, col_mail2 = st.columns(2)
        with col_mail1:
            st.link_button("🚀 Buka Langsung di Gmail Web", gmail_url, use_container_width=True)
        with col_mail2:
            st.link_button("📧 Buka di Aplikasi Email (HP / Desktop)", mailto_url, use_container_width=True)

        render_saweria_box()

# ==========================================
# TAB 4: WATERMARK PDF SKRIPSI (BERBAYAR VIA WHATSAPP)
# ==========================================
# --- Konfigurasi WhatsApp & Stok Kode ---
WA_PHONE = "6285240118570"
WA_MESSAGE = "Saya ingin membeli kode akses"
WA_PURCHASE_URL = f"https://wa.me/{WA_PHONE}?text={urllib.parse.quote(WA_MESSAGE)}"
LYNK_PRODUCT_URL = WA_PURCHASE_URL
EXCEL_KODE_PATH = os.path.join(os.path.dirname(__file__), "kode_akses_watermark.xlsx")
LYNK_ACCESS_SECRET = "SPI-LULUS-2026"  # Kunci rahasia untuk validasi token algoritmik NIM (legacy)
HMAC_SECRET_KEY = b"SPI-WATERMARK-UIN-JKT-2026-KEY"  # Kunci HMAC untuk validasi voucher kriptografis tanpa database

def generate_access_token(nim: str) -> str:
    """Menghasilkan kode akses unik berdasarkan NIM + kunci rahasia."""
    raw = f"{nim.strip()}-{LYNK_ACCESS_SECRET}"
    h = hashlib.sha256(raw.encode()).hexdigest()[:8].upper()
    return f"SPI-{h}"

def verify_voucher_code(token: str) -> bool:
    """Memverifikasi kode voucher format WM-XXXX-YYYY secara kriptografis tanpa database."""
    token_clean = token.strip().upper()
    parts = token_clean.split("-")
    if len(parts) != 3 or parts[0] != "WM":
        return False
    payload, sig = parts[1], parts[2]
    expected_sig = hmac.new(HMAC_SECRET_KEY, payload.encode("utf-8"), hashlib.sha256).hexdigest()[:4].upper()
    return hmac.compare_digest(sig, expected_sig)

def validate_access_token(nim: str, token: str) -> tuple[bool, str]:
    """Memvalidasi apakah kode akses cocok (via Voucher Kriptografis, Token NIM, atau Excel)."""
    token_clean = token.strip().upper()
    nim_clean = nim.strip()

    if not token_clean:
        return False, "⚠️ Kode Akses wajib diisi!"

    # 1. Verifikasi Kode Voucher Kriptografis (Format WM-XXXX-YYYY) - Bekerja tanpa database/server
    if verify_voucher_code(token_clean):
        return True, "🎉 Kode akses valid! Akses watermark PDF lengkap telah dibuka."

    # 2. Cek Kompatibilitas Token Algoritmik (NIM + Secret: SPI-XXXX)
    if nim_clean:
        expected = generate_access_token(nim_clean)
        if token_clean == expected:
            return True, "🎉 Kode akses valid (Algoritma NIM)! Akses watermark PDF lengkap telah dibuka."

    # 3. Cek stok kode di Excel lokal jika file tersedia
    if os.path.exists(EXCEL_KODE_PATH):
        try:
            import openpyxl
            wb = openpyxl.load_workbook(EXCEL_KODE_PATH)
            ws = wb.active
            for row in ws.iter_rows(min_row=2, max_col=4):
                code_val = str(row[1].value or "").strip().upper()
                status_val = str(row[2].value or "").strip().upper()

                if code_val == token_clean:
                    if status_val == "BELUM_DIPAKAI":
                        row[2].value = "SUDAH_DIPAKAI"
                        row[3].value = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                        wb.save(EXCEL_KODE_PATH)
                        return True, "🎉 Kode akses valid! Akses watermark PDF lengkap telah dibuka."
                    else:
                        return False, "❌ Kode akses ini sudah pernah digunakan sebelumnya."
        except Exception as e:
            pass

    return False, "❌ Kode akses tidak ditemukan atau tidak valid. Silakan periksa kembali."

with tab4:
    st.markdown(
        """
        <div style="background: linear-gradient(135deg, #667eea 0%, #764ba2 100%); padding: 20px 24px; border-radius: 12px; margin-bottom: 16px;">
            <h3 style="color: white; margin: 0 0 6px 0;">💧 Watermark PDF Skripsi Otomatis</h3>
            <p style="color: rgba(255,255,255,0.9); margin: 0; font-size: 14px;">
                Upload gambar watermark & PDF skripsi kamu, atur halaman mana yang perlu di-watermark, lalu unduh hasilnya.
            </p>
        </div>
        """,
        unsafe_allow_html=True
    )

    st.info("💡 **Cara Pakai:** Upload gambar watermark (logo/cap) dan PDF skripsi → Atur halaman → Lihat pratinjau 5 halaman gratis → Beli kode akses via Lynk.id → Unduh PDF lengkap ber-watermark.")

    # --- STEP 1: Upload Gambar Watermark ---
    st.markdown("### 🖼️ 1. Upload Gambar Watermark")
    st.caption("Format yang didukung: PNG, JPG, JPEG, WEBP. Gunakan gambar dengan latar transparan (PNG) untuk hasil terbaik.")
    wm_file = st.file_uploader(
        "Pilih file gambar watermark",
        type=["png", "jpg", "jpeg", "webp"],
        key="wm_upload",
        help="Upload logo UIN, logo fakultas, atau gambar kustom lainnya yang ingin dijadikan watermark."
    )

    if wm_file:
        wm_raw_bytes = wm_file.read()
        st.session_state["wm_raw"] = wm_raw_bytes

        col_wm_prev, col_wm_settings = st.columns([1, 2])
        with col_wm_prev:
            st.image(wm_raw_bytes, caption="Pratinjau gambar watermark", use_container_width=True)
        with col_wm_settings:
            wm_opacity = st.slider(
                "Tingkat Transparansi Watermark",
                min_value=5, max_value=50, value=15, step=5,
                format="%d%%",
                help="Semakin rendah nilainya, semakin transparan watermark-nya. Rekomendasi: 10-20%.",
                key="wm_opacity"
            )
            wm_scale = st.slider(
                "Ukuran Watermark (% dari halaman)",
                min_value=20, max_value=80, value=45, step=5,
                format="%d%%",
                help="Persentase lebar halaman yang akan ditempati watermark. Rekomendasi: 40-55%.",
                key="wm_scale"
            )
            wm_remove_bg = st.checkbox(
                "Hapus background putih dari gambar secara otomatis",
                value=False,
                help="Centang jika gambar watermark kamu berlatar putih (bukan transparan). Sistem akan menghapus background putihnya.",
                key="wm_rmbg"
            )
    elif "wm_raw" in st.session_state:
        wm_raw_bytes = st.session_state["wm_raw"]
        col_wm_prev, col_wm_settings = st.columns([1, 2])
        with col_wm_prev:
            st.image(wm_raw_bytes, caption="Pratinjau gambar watermark", use_container_width=True)
        with col_wm_settings:
            wm_opacity = st.slider(
                "Tingkat Transparansi Watermark",
                min_value=5, max_value=50, value=15, step=5,
                format="%d%%",
                help="Semakin rendah nilainya, semakin transparan watermark-nya. Rekomendasi: 10-20%.",
                key="wm_opacity"
            )
            wm_scale = st.slider(
                "Ukuran Watermark (% dari halaman)",
                min_value=20, max_value=80, value=45, step=5,
                format="%d%%",
                help="Persentase lebar halaman yang akan ditempati watermark. Rekomendasi: 40-55%.",
                key="wm_scale"
            )
            wm_remove_bg = st.checkbox(
                "Hapus background putih dari gambar secara otomatis",
                value=False,
                help="Centang jika gambar watermark kamu berlatar putih (bukan transparan). Sistem akan menghapus background putihnya.",
                key="wm_rmbg"
            )
    else:
        wm_raw_bytes = None

    st.markdown("---")

    # --- STEP 2: Upload PDF Skripsi ---
    st.markdown("### 📄 2. Upload PDF Skripsi")
    pdf_file = st.file_uploader(
        "Pilih file PDF skripsi yang akan di-watermark",
        type=["pdf"],
        key="wm_pdf_upload",
        help="Upload naskah skripsi dalam format PDF. Maksimal 200 MB."
    )

    if pdf_file:
        pdf_raw_bytes = pdf_file.read()
        st.session_state["wm_pdf_raw"] = pdf_raw_bytes

        try:
            tmp_doc = pymupdf.open(stream=pdf_raw_bytes, filetype="pdf")
            total_pages = len(tmp_doc)
            tmp_doc.close()
            st.session_state["wm_total_pages"] = total_pages
            st.success(f"✅ PDF berhasil dimuat: **{total_pages} halaman** ({len(pdf_raw_bytes) / 1024 / 1024:.1f} MB)")
        except Exception as e:
            st.error(f"❌ Gagal membaca file PDF: {e}")
            total_pages = 0
    elif "wm_pdf_raw" in st.session_state:
        pdf_raw_bytes = st.session_state["wm_pdf_raw"]
        total_pages = st.session_state.get("wm_total_pages", 0)
        if total_pages > 0:
            st.success(f"✅ PDF dimuat: **{total_pages} halaman** ({len(pdf_raw_bytes) / 1024 / 1024:.1f} MB)")
    else:
        pdf_raw_bytes = None
        total_pages = 0

    # --- STEP 3: Pengaturan Halaman ---
    if total_pages > 0 and wm_raw_bytes:
        st.markdown("---")
        st.markdown("### ⚙️ 3. Pengaturan Halaman yang Di-watermark")

        page_mode = st.radio(
            "Pilih mode halaman:",
            ["Semua Halaman", "Semua Kecuali Cover (Halaman 1)", "Kustom (Pilih Sendiri)"],
            horizontal=True,
            key="wm_page_mode"
        )

        skip_str = ""
        include_str = ""
        if page_mode == "Kustom (Pilih Sendiri)":
            col_inc, col_skip = st.columns(2)
            with col_inc:
                include_str = st.text_input(
                    f"Halaman yang DI-watermark (1-{total_pages})",
                    placeholder="Contoh: 1-5, 10, 15-20",
                    help="Kosongkan untuk memilih semua halaman, lalu gunakan kolom 'lewati' untuk mengecualikan.",
                    key="wm_include"
                )
            with col_skip:
                skip_str = st.text_input(
                    "Halaman yang DILEWATI (tidak di-watermark)",
                    placeholder="Contoh: 1, 2, 100",
                    help="Halaman-halaman ini tidak akan diberi watermark.",
                    key="wm_skip"
                )

        clean_mode = page_mode.split(" (")[0] if "(" in page_mode else page_mode
        selected_pages = parse_page_selection(total_pages, clean_mode, skip_str, include_str)
        wm_count = len(selected_pages)
        skip_count = total_pages - wm_count

        st.markdown(
            f"""
            <div style="background-color: #f0f9ff; border: 1px solid #bae6fd; border-radius: 8px; padding: 12px 16px; margin: 8px 0;">
                <span style="font-size: 14px;">
                    📊 <b>{wm_count}</b> halaman akan di-watermark &nbsp;|&nbsp;
                    ⏭️ <b>{skip_count}</b> halaman dilewati &nbsp;|&nbsp;
                    📑 Total: <b>{total_pages}</b> halaman
                </span>
            </div>
            """,
            unsafe_allow_html=True
        )

        st.markdown("---")

        # --- STEP 4: Generate Pratinjau Gratis (5 Halaman) ---
        st.markdown("### 👁️ 4. Pratinjau Gratis (5 Halaman Pertama)")

        if st.button("🔍 Lihat Pratinjau 5 Halaman", use_container_width=True, key="wm_preview_btn"):
            with st.spinner("Sedang memproses pratinjau watermark..."):
                try:
                    prep_wm = prepare_watermark_image(
                        wm_raw_bytes,
                        opacity=st.session_state.get("wm_opacity", 15) / 100,
                        remove_white_bg=st.session_state.get("wm_rmbg", False)
                    )
                    _, sample_imgs = apply_watermark_to_pdf(
                        pdf_raw_bytes,
                        prep_wm,
                        selected_pages,
                        scale=st.session_state.get("wm_scale", 45) / 100,
                        sample_only=5
                    )
                    st.session_state["wm_samples"] = sample_imgs
                except Exception as e:
                    st.error(f"❌ Gagal memproses pratinjau: {e}")

        if "wm_samples" in st.session_state and st.session_state["wm_samples"]:
            samples = st.session_state["wm_samples"]
            st.caption(f"Menampilkan {len(samples)} halaman pertama sebagai pratinjau:")

            # Tampilkan 2-3 kolom per baris
            cols_per_row = 3 if len(samples) >= 3 else len(samples)
            for row_start in range(0, len(samples), cols_per_row):
                row_imgs = samples[row_start:row_start + cols_per_row]
                cols = st.columns(len(row_imgs))
                for col_idx, img_bytes in enumerate(row_imgs):
                    with cols[col_idx]:
                        page_num = row_start + col_idx + 1
                        is_watermarked = (page_num - 1) in selected_pages
                        label = f"Hal. {page_num}" + (" 💧" if is_watermarked else " ⏭️")
                        st.image(img_bytes, caption=label, use_container_width=True)

            st.markdown("---")

        # --- STEP 5: Akses Berbayar via Lynk.id ---
        st.markdown("### 🔐 5. Unduh PDF Lengkap Ber-watermark")

        # Cek apakah sudah unlock via query param (redirect dari Lynk.id)
        params = st.query_params
        if params.get("wm_akses") == "sukses":
            st.session_state["wm_unlocked"] = True

        if st.session_state.get("wm_unlocked", False):
            # SUDAH UNLOCK — Tombol proses & download penuh
            st.success("🔓 **Akses terbuka!** Kamu bisa memproses dan mengunduh PDF lengkap ber-watermark.")

            if st.button("🚀 Proses Seluruh PDF Ber-watermark", use_container_width=True, key="wm_full_btn", type="primary"):
                with st.spinner(f"Sedang memproses {wm_count} halaman... Mohon tunggu sebentar."):
                    try:
                        prep_wm = prepare_watermark_image(
                            wm_raw_bytes,
                            opacity=st.session_state.get("wm_opacity", 15) / 100,
                            remove_white_bg=st.session_state.get("wm_rmbg", False)
                        )
                        full_pdf, _ = apply_watermark_to_pdf(
                            pdf_raw_bytes,
                            prep_wm,
                            selected_pages,
                            scale=st.session_state.get("wm_scale", 45) / 100,
                            sample_only=None
                        )
                        st.session_state["wm_full_pdf"] = full_pdf
                    except Exception as e:
                        st.error(f"❌ Gagal memproses PDF: {e}")

            if "wm_full_pdf" in st.session_state:
                full_data = st.session_state["wm_full_pdf"]
                st.success(f"✅ PDF ber-watermark berhasil diproses! Ukuran: {len(full_data) / 1024 / 1024:.1f} MB")
                btn_dl_full = st.download_button(
                    label=f"📥 Unduh PDF Ber-watermark ({total_pages} halaman)",
                    data=full_data,
                    file_name="Skripsi_Watermarked.pdf",
                    mime="application/pdf",
                    use_container_width=True,
                    key="wm_dl_full"
                )
                if btn_dl_full:
                    st.toast("🎉 PDF ber-watermark berhasil diunduh!")
                    st.balloons()

        else:
            # BELUM UNLOCK — Tampilkan paywall
            st.markdown(
                """
                <div style="background: linear-gradient(135deg, #fef3c7, #fde68a); border: 2px solid #f59e0b; border-radius: 12px; padding: 20px; margin: 12px 0;">
                    <h4 style="margin: 0 0 8px 0; color: #92400e;">🔒 Fitur Premium — Watermark PDF Lengkap</h4>
                    <p style="margin: 0 0 12px 0; color: #78350f; font-size: 14px; line-height: 1.6;">
                        Pratinjau 5 halaman pertama <b>gratis</b> untuk memastikan hasil watermark sesuai keinginanmu.<br>
                        Untuk memproses dan mengunduh <b>seluruh halaman</b> PDF skripsi ber-watermark, silakan beli kode akses melalui <b>WhatsApp Admin</b> (QRIS, GoPay, OVO, DANA, ShopeePay, Transfer Bank).<br><br>
                        💰 <b>Harga: Rp5.000</b> (sekali bayar per kode akses)
                    </p>
                </div>
                """,
                unsafe_allow_html=True
            )

            st.link_button(
                "📲 Beli Kode Akses via WhatsApp (Rp5.000)",
                WA_PURCHASE_URL,
                use_container_width=True
            )

            st.markdown("---")
            st.markdown("##### Sudah punya kode akses? Masukkan di bawah:")

            col_nim_wm, col_token = st.columns(2)
            with col_nim_wm:
                wm_nim = st.text_input(
                    "NIM Kamu (Opsional)",
                    placeholder="Contoh: 11200210000088",
                    key="wm_nim",
                    help="NIM kamu (opsional jika menggunakan kode unik WhatsApp)."
                )
            with col_token:
                wm_token = st.text_input(
                    "Kode Akses",
                    placeholder="Contoh: WM-ABCD-1234",
                    key="wm_token",
                    help="Kode akses yang kamu terima dari Admin WhatsApp."
                )

            if st.button("🔓 Validasi & Buka Akses", use_container_width=True, key="wm_validate_btn"):
                is_valid, msg = validate_access_token(wm_nim, wm_token)
                if is_valid:
                    st.session_state["wm_unlocked"] = True
                    st.success(msg)
                    st.rerun()
                else:
                    st.error(msg)

    elif not wm_raw_bytes and not pdf_raw_bytes:
        st.warning("⬆️ Silakan upload **gambar watermark** dan **file PDF skripsi** terlebih dahulu di atas.")
    elif not wm_raw_bytes:
        st.warning("⬆️ Silakan upload **gambar watermark** terlebih dahulu.")
    elif total_pages == 0:
        st.warning("⬆️ Silakan upload **file PDF skripsi** terlebih dahulu.")
