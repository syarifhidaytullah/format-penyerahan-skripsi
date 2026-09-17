# Portal Layanan Berkas Skripsi SPI UIN Jakarta

Aplikasi web sederhana berbasis Python & Streamlit untuk mengotomatisasi pembuatan formulir resmi administrasi skripsi Program Studi Sejarah dan Peradaban Islam (SPI), Fakultas Adab dan Humaniora, UIN Syarif Hidayatullah Jakarta.

## Fitur 5 Layanan Lengkap:
1. **📄 Tanda Bukti Penyerahan Skripsi**
   - Mengisi otomatis data Mahasiswa, NIM, Prodi, Tanggal Sidang, dan Judul Skripsi.
   - Deteksi pintar format nomor Dosen: 18 digit = NIP (PNS/ASN), 10 digit = NIDN (Non-PNS).
   - Padding 1 baris antara Judul dan Tabel dengan layout presisi 1 lembar A4.
   - Pilihan unduh versi PDF (siap cetak) atau Word (.docx).

2. **📝 Formulir Pendaftaran Ujian Skripsi**
   - Mengisi otomatis data lengkap pendaftaran munaqasyah: Tempat/Tgl Lahir, Asal SLTA, Alamat, IPK, Judul, No HP/Email.
   - Pilihan Semester (Ganjil/Genap) & Tahun Akademik.
   - Blok tanda tangan pemohon terisi otomatis tanpa catatan bintang berlebih.
   - Presisi 1 lembar A4 dalam format PDF dan Word (.docx).

3. **📋 Formulir Pendaftaran Sidang (Persyaratan Map Kuning)**
   - Mengisi data Mahasiswa, NIM, Prodi, dan Tanggal Surat.
   - Memuat lengkap 11 poin checklist berkas persyaratan ujian skripsi map kuning.
   - Menampilkan kop surat resmi dan tanda tangan Kabag TU FAH UIN Jakarta.
   - Presisi 1 lembar A4 dalam format PDF dan Word (.docx).

4. **📌 Panduan & Checklist Persyaratan (Pra-Sidang & Pasca-Sidang)**
   - Checklist interaktif 12 berkas persyaratan ujian skripsi map kuning.
   - Alur persyaratan pengurusan BAP dan Transkrip Nilai (revisi ber-watermark, submit jurnal Socio Historica).
   - Informasi alamat email resmi Program Studi SPI.

5. **✉️ Template Email Siap Salin & Buka di Gmail**
   - Generator otomatis format email resmi ke prodi sesuai template infografis resmi.
   - Skenario 1: Permohonan Pendaftaran Ujian Proposal / Skripsi.
   - Skenario 2: Permohonan Berita Acara Sidang Skripsi (BAP).
   - Fitur salin 1-klik dan tombol langsung `Buka di Gmail` / `Aplikasi Email` dengan subjek dan isi yang sudah terisi otomatis.

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
