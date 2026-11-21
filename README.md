# Buku Induk Siswa

![Flask](https://img.shields.io/badge/Flask-2f81f7?style=flat) ![Python](https://img.shields.io/badge/Python-2f81f7?style=flat) ![SQLite](https://img.shields.io/badge/SQLite-2f81f7?style=flat) ![openpyxl](https://img.shields.io/badge/openpyxl-2f81f7?style=flat) ![Jinja2](https://img.shields.io/badge/Jinja2-2f81f7?style=flat)

> Aplikasi pengganti buku induk manual sekolah yang berjalan lokal tanpa server terpisah.

## Tentang proyek

Aplikasi pengganti buku induk manual sekolah yang berjalan lokal tanpa server terpisah. Dokumentasi ini dibuat sebagai panduan singkat untuk memahami ruang lingkup proyek, menjalankannya secara lokal, dan membaca batasan penggunaannya.

## Fitur utama

- Import massal Excel dengan upsert berbasis NIS
- Data identitas dan nilai per semester
- Dashboard statistik dan pencarian siswa
- Cetak atau ekspor laporan per siswa
- Penyimpanan lokal berbasis SQLite

## Teknologi

- Flask
- Python
- SQLite
- openpyxl
- Jinja2

## Menjalankan secara lokal

Prasyarat dan konfigurasi dapat berbeda menurut sistem operasi. Contoh langkah utama:

```bash
python app.py
```

## Struktur dan pengembangan

Mulai dari entry point dan folder konfigurasi proyek. Pisahkan credential lokal dari source code, tambahkan test untuk perubahan penting, dan gunakan environment variable untuk konfigurasi deployment.

## Keamanan

- Jangan commit API key, password, token, session ID, atau data pribadi.
- Gunakan file environment lokal dan contoh konfigurasi tanpa nilai rahasia.
- Uji perubahan pada environment development sebelum deployment.

## Status

Proyek ini merupakan bagian dari portofolio **Hilman Arifin Ilham** dan dapat dikembangkan sesuai kebutuhan penggunaan.

> **Catatan:** File runtime seperti `data/secret.key` tidak termasuk repository publik. Buat konfigurasi lokal sesuai kebutuhan sebelum menjalankan aplikasi.

## Lisensi

Belum ada lisensi open-source yang dinyatakan di repository ini. Hubungi pemilik repository sebelum menggunakan ulang atau mendistribusikan kode.

---

Dibuat oleh [Hilman Arifin Ilham](https://github.com/hilman-devop).
