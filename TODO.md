## Rencana Kerja (Future Work)

### Prioritas Utama
1. [ X ] **Validasi Input Kritis**:
    - [ X ] Terapkan `QDoubleValidator` atau `QIntValidator` untuk semua field numerik (misalnya, Temperatur, Nilai Min/Max, Deviasi) untuk mencegah input non-angka yang menyebabkan error saat pembuatan dokumen. Ini adalah item paling penting untuk stabilitas aplikasi.

### Peningkatan & Kelayakan Pengguna
2. [X] **Peningkatan Penanganan Error & Umpan Balik**:
    - [X] Tambahkan pemeriksaan untuk file template (`New_Template2.docx`) **sebelum** mencoba membuat dokumen. Tampilkan pesan error yang jelas jika file tidak ditemukan, daripada membiarkan aplikasi crash.
    - [X] Saat memuat proyek, jika path file gambar tidak valid, tampilkan pesan yang informatif di `QLineEdit` (misalnya, "File tidak ditemukan") daripada hanya menampilkan path yang rusak.

3. [X] **Distribusi Aplikasi yang Mudah**:
    - [X] Siapkan skrip atau konfigurasi untuk mengemas aplikasi menjadi satu file `.exe` menggunakan **PyInstaller** atau cx_Freeze agar pengguna non-teknis dapat menjalankannya tanpa menginstal Python.
    - [X] Perbarui `README.md` dengan instruksi cara menjalankan file `.exe` tersebut.

---

## Selesai (Completed) - 13 Nov 2025

- [x] **Peningkatan Penanganan Error & Umpan Balik**: Menambahkan pemeriksaan untuk file template yang hilang dan path gambar yang tidak valid.
- [x] **Distribusi Aplikasi yang Mudah**: Menambahkan instruksi untuk PyInstaller dan memperbarui dokumentasi (`README.md`, `usermanual.md`).
- [x] **Manajemen File Sementara**: Memperbaiki masalah penumpukan file sementara saat aplikasi ditutup.
- [x] **Fitur Simpan/Muat Proyek**: Mengimplementasikan logika penyimpanan yang benar sehingga gambar (termasuk screenshot) disimpan bersama proyek dan dapat dimuat kembali.
- [x] **Path Template yang Robust**: Memperbaiki cara aplikasi menemukan file template agar tidak bergantung pada direktori kerja saat aplikasi dijalankan.