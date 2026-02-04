# HR Analytics Tools

**Sistem Otomatisasi Audit Absensi & Visualisasi Kinerja Karyawan (Plant Area)**

![Python](https://img.shields.io/badge/Python-3.10%2B-blue?style=for-the-badge&logo=python&logoColor=white)
![PyQt6](https://img.shields.io/badge/GUI-PyQt6-green?style=for-the-badge&logo=qt&logoColor=white)
![Pandas](https://img.shields.io/badge/Data-Pandas-150458?style=for-the-badge&logo=pandas&logoColor=white)

## Deskripsi
**HR Analytics Tools** adalah aplikasi desktop untuk mengotomatisasi audit data absensi karyawan pabrik (*Plant*) di **PT Sumber Indah Perkasa**. Aplikasi ini mengatasi masalah inkonsistensi nama karyawan, mendeteksi anomali (telat/pulang awal) dengan logika *State Machine*, dan menghasilkan laporan visual (PDF/Excel) secara otomatis.

---

## Fitur Utama
1.  **Smart Identity Sync:** Memperbaiki nama karyawan beda ejaan secara otomatis (Key-Value Mapping).
2.  **Deteksi Anomali:** Menghitung keterlambatan, pulang awal, dan pelanggaran istirahat secara presisi.
3.  **Laporan Visual:** Generate grafik *Stacked Bar* & *Heatmap* untuk manajemen.
4.  **Cek Lembur:** Modul *live preview* untuk validasi jam lembur.

---
## Tangkapan Layar (Screenshots)
### 1. Menu Utama (Dashboard)
Navigasi utama untuk memilih modul analisis.
![Menu Utama](assets/HU.png)
### 2. Modul 1: Analisis Absensi Plant
![Modul Analisis](assets/M1.png)
### 3. Modul 2: Pemeriksa Data Lembur
![Modul Analisis](assets/M2.png)


## Cara Menggunakan (Quick Start)

Pastikan Python 3.10+ sudah terinstal. Buka terminal/CMD, lalu jalankan perintah berikut secara berurutan:

```bash
# 1. Clone repository & masuk ke folder
git clone [https://github.com/kajo737/NAMA_REPO.git](https://github.com/kajo737/PyQt-HR-Analytics-Tools__.git)
cd PyQt-HR-Analytics-Tools__

# 2. Install semua library (hanya sekali di awal)
pip install -r requirements.txt

# 3. Jalankan aplikasi
python Aplikasi_HR.py
