# Portal Layanan Berkas Skripsi SPI UIN Jakarta

Aplikasi web sederhana berbasis Python & Streamlit untuk mengotomatisasi pembuatan formulir resmi administrasi skripsi Program Studi Sejarah dan Peradaban Islam (SPI), Fakultas Adab dan Humaniora, UIN Syarif Hidayatullah Jakarta.

## Fitur 8 Layanan Lengkap:
1. **📌 Panduan & Checklist Persyaratan (Pra-Sidang & Pasca-Sidang)**
   - Checklist interaktif 12 berkas persyaratan ujian skripsi map kuning.
   - Alur persyaratan pengurusan BAP dan Transkrip Nilai (revisi ber-watermark, submit jurnal Socio Historica).
   - Informasi alamat email resmi Program Studi SPI.

2. **📄 Tanda Bukti Penyerahan Skripsi**
   - Mengisi otomatis data Mahasiswa, NIM, Prodi, Tanggal Sidang, dan Judul Skripsi.
   - Deteksi pintar format nomor Dosen: 18 digit = NIP (PNS/ASN), 10 digit = NIDN (Non-PNS).
   - Padding 1 baris antara Judul dan Tabel dengan layout presisi 1 lembar A4.
   - Pilihan unduh versi PDF (siap cetak) atau Word (.docx).

3. **📝 Formulir Pendaftaran Ujian Skripsi**
   - Mengisi otomatis data lengkap pendaftaran munaqasyah: Tempat/Tgl Lahir, Asal SLTA, Alamat, IPK, Judul, No HP/Email.
   - Pilihan Semester (Ganjil/Genap) & Tahun Akademik.
   - Blok tanda tangan pemohon terisi otomatis tanpa catatan bintang berlebih.
   - Presisi 1 lembar A4 dalam format PDF dan Word (.docx).

4. **📋 Formulir Pendaftaran Sidang (Persyaratan Map Kuning)**
   - Mengisi data Mahasiswa, NIM, Prodi, dan Tanggal Surat.
   - Memuat lengkap 11 poin checklist berkas persyaratan ujian skripsi map kuning.
   - Menampilkan kop surat resmi dan tanda tangan Kabag TU FAH UIN Jakarta.
   - Presisi 1 lembar A4 dalam format PDF dan Word (.docx).

5. **📖 Lembar Bimbingan Skripsi**
   - Menghasilkan Lembar Bimbingan Skripsi resmi prodi ber-kop resmi Fakultas Adab dan Humaniora.
   - Memuat tabel riwayat bimbingan 8 baris terstruktur (syarat minimal 6 kali konsultasi pembimbing).
   - Dilengkapi blok tanda tangan berdampingan: Ketua Program Studi (kiri) dan Dosen Pembimbing Skripsi (kanan).
   - Presisi tepat 1 lembar A4 dalam format PDF dan Word (.docx).

6. **✍️ Lembar Pernyataan Skripsi (Bebas Plagiasi)**
   - Format baku surat pernyataan orisinalitas karya ilmiah dan bebas plagiarisme sesuai pedoman akademik UIN Jakarta.
   - Memuat data Mahasiswa, NIM, Prodi, 3 butir pernyataan hukum/akademik, dan blok tanda tangan pemohon.
   - Opsi panduan posisi tempel Materai 10.000 jika diwajibkan oleh fakultas/perpustakaan.
   - Presisi tepat 1 lembar A4 dalam format PDF dan Word (.docx).

7. **🏛️ Surat Pernyataan Izin Publikasi Repository Perpustakaan**
   - Surat resmi izin publikasi karya ilmiah di Repository Perpustakaan UIN Syarif Hidayatullah Jakarta.
   - Dilengkapi kop resmi Fakultas Adab dan Humaniora dan ruang materai 10.000.
   - Presisi tepat 1 lembar A4 dalam format PDF dan Word (.docx).

8. **✉️ Template Email Siap Salin & Buka di Gmail**
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
