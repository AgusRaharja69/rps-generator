# 📘 RPS Generator

[![Build Status](https://img.shields.io/badge/build-passing-brightgreen)](https://github.com/your-repo)  
*Aplikasi ini digunakan untuk membuat dokumen **RPS (Rencana Pembelajaran Semester), RTM, Kontrak, Rubrik, Portofolio Penilaian, dan Nilai Format OBE** secara otomatis menggunakan Python dan Google Sheets API. Program ini mempermudah proses pembuatan RPS dengan mengotomatiskan pengisian data dari spreadsheet.*

---

## 📋 Overview

| Fitur                  | Deskripsi                              |
|-------------------------|----------------------------------------|
| Otomatisasi RPS         | Membuat RPS berdasarkan data spreadsheet |
| Integrasi Google Sheets | Menggunakan API untuk pembaruan data   |
| Penyesuaian Fleksibel   | Mendukung modifikasi kode              |

## 📦 Persyaratan

- Python 3.x
- Akses ke Google Sheets API
- Koneksi internet untuk autentikasi dan pembaruan data

## 🚀 Langkah Penggunaan

### 1. Instalasi Dependensi
Pastikan Anda telah menginstal Python dan `pip`. Instal dependensi yang diperlukan dengan menjalankan perintah berikut di terminal:

```bash
pip install -r requirements.txt
```

Jika file `requirements.txt` belum ada, buat dengan daftar pustaka seperti `gspread`, `oauth2client`, atau lainnya yang digunakan.

### 2. Salin Template Spreadsheet
Salin spreadsheet template berikut sebagai dasar untuk RPS Anda:
- 👉 [Template Spreadsheet](https://docs.google.com/spreadsheets/d/1-H_m3vrxRV7YrvQ5w-5McfOt_NNN54-_19sb2JZwfqE/edit?gid=1757757975#gid=1757757975)

Spreadsheet ini akan menjadi tempat penyimpanan dan pengolahan data RPS secara otomatis.

### 3. Lengkapi Data Mata Kuliah
Buka sheet bernama **Robotika (K1)** LALU COPY SHEET INI UNTUK MATA KULIAH ANDA, pada spreadsheet yang telah disalin

Pastikan data diisi dengan benar untuk memastikan hasil yang akurat DAN SESUAI DENGAN FORMAT.

### 4. Berikan Akses ke Google API Service
Berikan akses editor ke akun service berikut agar aplikasi dapat mengakses spreadsheet:

```
python-rps-api@rps-generator.iam.gserviceaccount.com
```

**Langkah-langkah:**
- Buka Google Sheets.
- Bagikan spreadsheet dengan akun di atas sebagai "Editor".

### 5. Siapkan File `credentials.json` (OPTIONAL)
File `credentials.json` berisi kredensial API dari Google Cloud Console. Jika belum ada BISA CONTACT DEVELOPER atau:
- Buat proyek di [Google Cloud Console](https://console.cloud.google.com/).
- Aktifkan Google Sheets API.
- Buat service account dan unduh file JSON.
- Simpan file tersebut sebagai `credentials.json` di direktori proyek ini.

### 6. Jalankan Program
Setelah semua langkah di atas selesai, jalankan aplikasi dengan perintah:

```bash
python main.py
```

### ✅ Hasil
Program akan secara otomatis menghasilkan dokumen RPS berdasarkan data yang telah Anda isi. Hasilnya akan diperbarui di spreadsheet yang sama.

---

## 🎉 Selesai!
Selamat! 🎓 RPS Anda telah berhasil dibuat secara otomatis. Periksa spreadsheet untuk memastikan semua data terisi dengan benar.

---

## 🛠️ Catatan Penting
- Pastikan koneksi internet stabil saat menjalankan program.
- Jika terjadi error, periksa log atau pesan kesalahan yang muncul di terminal.
- Untuk penyesuaian tambahan (misalnya, format warna atau penomoran), modifikasi kode di `main.py` sesuai kebutuhan.

## 🤝 Kontribusi
Kontribusi diterima! Silakan buka issue atau pull request di repositori jika ada saran atau perbaikan.

## 📄 Lisensi
