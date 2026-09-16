# Portal Layanan Berkas Skripsi SPI UIN Jakarta

Aplikasi web sederhana berbasis Python & Streamlit untuk mengotomatisasi pembuatan formulir resmi administrasi skripsi Program Studi Sejarah dan Peradaban Islam (SPI), Fakultas Adab dan Humaniora, UIN Syarif Hidayatullah Jakarta.

## Fitur Berkas:
1. **📄 Tanda Bukti Penyerahan Skripsi**
   - Mengisi otomatis data Mahasiswa, NIM, Prodi, Tanggal Sidang, dan Judul Skripsi.
   - Deteksi pintar format nomor Dosen: 18 digit = NIP (PNS/ASN), 10 digit = NIDN (Non-PNS).
   - Padding 1 baris antara Judul dan Tabel dengan layout presisi 1 lembar A4.
   - Pilihan unduh versi PDF (siap cetak) atau Word (.docx).

2. **📝 Formulir Pendaftaran Ujian Skripsi**
   - Mengisi otomatis data lengkap pendaftaran munaqasyah: TTL, Asal SLTA, Alamat, IPK, Judul, No HP/Email.
   - Pilihan Semester (Ganjil/Genap) & Tahun Akademik.
   - Tanda tangan pemohon terisi otomatis.
   - Presisi 1 lembar A4 dalam format PDF dan Word (.docx).

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
