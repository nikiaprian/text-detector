# Ekstrak Teks Biru dari PDF

Halaman web sederhana (HTML + Python Flask) untuk mengunggah PDF, membaca **hanya teks yang berwarna biru**, lalu menghasilkan PDF baru yang berisi teks biru saja.

## Cara menjalankan

1. Pasang dependensi:
   ```bash
   pip install -r requirements.txt
   ```

2. Jalankan server:
   ```bash
   python app.py
   ```

3. Buka di browser: **http://localhost:5000**

4. Pilih file PDF → klik **Ekstrak teks biru → unduh PDF**. File `teks_biru.pdf` akan terunduh.

## Catatan

- **Warna biru** di sini: komponen biru (B) lebih besar dari merah (R) dan hijau (G), dan B ≥ 80. Jika PDF Anda memakai nuansa biru lain, bisa disesuaikan di `app.py` (fungsi `is_blue_color`).
- Beberapa PDF (misalnya hasil export dari browser) bisa tidak menyimpan informasi warna teks; dalam kasus itu tidak ada teks biru yang terdeteksi.
