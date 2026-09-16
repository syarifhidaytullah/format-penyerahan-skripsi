# Generator Formulir Tanda Bukti Penyerahan Skripsi - SPI UIN Jakarta

Aplikasi web sederhana berbasis Python & Streamlit untuk mengotomatisasi pembuatan formulir resmi Tanda Bukti Penyerahan Skripsi Program Studi Sejarah dan Peradaban Islam (SPI), Fakultas Adab dan Humaniora, UIN Syarif Hidayatullah Jakarta.

## Fitur
- Pengisian data Nama, NIM, Program Studi, Tanggal Sidang, dan Judul Skripsi.
- Format penanggalan surat otomatis dalam Bahasa Indonesia (misal: "Jakarta, 16 September 2026").
- Mengubah warna font tanggal/sidang yang sebelumnya merah di template resmi menjadi hitam otomatis.
- Opsional pengisian Dosen Pembimbing dan Dosen Penguji.
- Langsung menghasilkan dan mengunduh file Microsoft Word (.docx) tanpa mengubah margin, tabel, kop surat, atau indentasi aslinya.

## Menjalankan di Lokal

```bash
# Buat venv dan install dependensi
pip install -r requirements.txt

# Jalankan aplikasi Streamlit
streamlit run app.py
```

## Deploy ke Streamlit Community Cloud (Gratis & Online 24 Jam)
1. Push repository ini ke GitHub.
2. Buka [share.streamlit.io](https://share.streamlit.io).
3. Sambungkan repository GitHub ini, pilih branch `main`, dan set Main file path ke `app.py`.
4. Klik **Deploy**.
