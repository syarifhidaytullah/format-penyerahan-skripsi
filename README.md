# Portal Layanan Berkas Skripsi SPI UIN Jakarta

Aplikasi web sederhana berbasis Python & Streamlit untuk mengotomatisasi pembuatan formulir resmi administrasi skripsi Program Studi Sejarah dan Peradaban Islam (SPI), Fakultas Adab dan Humaniora, UIN Syarif Hidayatullah Jakarta.

## Struktur Navigasi Ramah HP (3 Tab Alur Akademik):

Aplikasi didesain khusus agar nyaman dibuka di layar HP (smartphone) tanpa perlu menggeser deretan tab panjang. Semua kebutuhan dikelompokkan ke dalam **3 Tab Alur Utama**:

### 1. 📌 Tab 1: Panduan & Alur Berkas (Default View)
- Checklist interaktif 12 berkas persyaratan ujian skripsi (map kuning).
- Alur persyaratan pengurusan BAP dan Transkrip Nilai pasca-sidang.
- Tautan langsung ke Jurnal *Socio Historica* UIN Jakarta dan alamat email resmi prodi SPI.

### 2. 📝 Tab 2: Berkas Pendaftaran (Pra-Sidang)
Tersedia sub-menu pilihan cepat untuk membuat dokumen persyaratan sidang:
- **Formulir Permohonan Ujian Skripsi (TU):** Isian identitas, IPK, asal SLTA, semester/tahun akademik.
- **Formulir Pendaftaran Sidang (Map Kuning):** Memuat 11 checklist berkas dan tanda tangan Kabag TU.
- **Lembar Bimbingan Skripsi:** Memuat kop resmi FAH, tabel riwayat konsultasi 8 baris (min. 6 kali), dan tanda tangan berdampingan Kaprodi & Pembimbing.
- **Lembar Pernyataan Skripsi:** Surat pernyataan bebas plagiasi & orisinalitas karya ilmiah 3 paragraf presisi 1 lembar A4 (dengan opsi panduan Materai 10.000).

### 3. 🎓 Tab 3: Berkas Kelulusan (Pasca-Sidang)
Tersedia sub-menu pilihan cepat untuk administrasi kelulusan:
- **Tanda Bukti Penyerahan Skripsi:** Bukti serah terima naskah revisi ke dosen penguji, pembimbing, fakultas, dan perpus.
- **Surat Pernyataan Izin Publikasi Repository:** Izin unggah karya ke perpustakaan UIN Jakarta bermaterai 10.000.
- **Template Email Siap Salin:** Generator email resmi prodi (skenario pendaftaran sidang & permohonan BAP) + tombol langsung buka di Gmail / HP mail client.

---

## Fitur Cerdas Lainnya:
- **Sinkronisasi Data Otomatis:** Input Nama, NIM, Prodi, Judul, atau Nama Dosen di satu formulir akan otomatis mengisi formulir lainnya sehingga mahasiswa tidak perlu mengetik ulang berulang kali.
- **Presisi 1 Lembar A4:** Semua dokumen dikalibrasi agar pas tepat 1 halaman A4 saat diunduh via PDF maupun Word (.docx).
- **Deteksi Otomatis NIP vs NIDN:** 18 digit angka dikenali sebagai NIP (ASN), 10 digit angka dikenali sebagai NIDN.

## Menjalankan di Lokal:
```bash
# Install dependensi
pip install -r requirements.txt

# Jalankan server Streamlit
streamlit run app.py
```

## Deploy ke Streamlit Community Cloud (Gratis 24 Jam):
1. Buka [share.streamlit.io](https://share.streamlit.io) dan login dengan akun GitHub.
2. Sambungkan repo `syarifhidaytullah/format-penyerahan-skripsi`.
3. Main file path: `app.py`.
4. Klik **Deploy**.
