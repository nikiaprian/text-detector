"""
Varian backend Flask (app2):
- Reuse seluruh logika ekstraksi di `app.py`
- Output tabel utama langsung memakai `raw_data` (18 kolom) yang dikembalikan
  oleh `build_table_with_header_from_pdf`, supaya tampilan tabel konsisten
  dengan raw data yang diambil dari PDF.

Catatan:
- Import fungsi-fungsi pemrosesan PDF dari `app.py` untuk menghindari
  duplikasi 3.000+ baris kode.
"""

import os
import tempfile
from io import BytesIO

from flask import Flask, request, send_file, render_template, jsonify

from app import (  # type: ignore
    build_table_with_header_from_pdf,
    extract_blue_spans_with_bbox,
    create_pdf_raw_blue_one_per_line,
    create_pdf_from_table,
    _looks_like_percentage_value,
    _looks_like_large_number,
)


app = Flask(__name__)
app.config["MAX_CONTENT_LENGTH"] = 16 * 1024 * 1024  # 16 MB max upload


def _column_index_by_name(header_row: list[str], candidates: tuple[str, ...]) -> int:
    """Cari indeks kolom di header berdasarkan beberapa kandidat nama (case-insensitive)."""
    if not header_row:
        return -1
    header_lower = [str(h).strip().lower() for h in header_row]
    for name in candidates:
        name_l = name.strip().lower()
        for idx, col in enumerate(header_lower):
            if col == name_l:
                return idx
    for name in candidates:
        name_l = name.strip().lower()
        for idx, col in enumerate(header_lower):
            if name_l in col:
                return idx
    return -1


def _fill_persentase_from_same_row(table: list[list[str]]) -> None:
    """
    Isi Persentase (1) dan Persentase (2) yang kosong dari nilai persen lain di baris yang sama.
    Logika generik (tanpa hardcode No): kumpulkan semua nilai mirip persen di kolom 11-17;
    jika ada 2 nilai, isi (1) dan (2); jika (2) kosong dan Perubahan berisi persen, pindahkan ke (2).
    """
    if not table or len(table) < 2:
        return

    idx_pct1 = 13
    idx_pct2 = 16
    idx_perubahan = 17
    numeric_cols = (11, 12, 13, 14, 15, 16, 17)

    for row in table[1:]:
        while len(row) <= idx_perubahan:
            row.append("-")
        pct1 = str(row[idx_pct1]).strip() if idx_pct1 < len(row) else ""
        pct2 = str(row[idx_pct2]).strip() if idx_pct2 < len(row) else ""
        if pct1 and pct1 != "-" and pct2 and pct2 != "-":
            continue

        # Kumpulkan nilai persen di baris (urut indeks kolom) beserta indeksnya
        ordered_pct: list[tuple[int, str]] = []
        seen_val: set[str] = set()
        for j in numeric_cols:
            if j >= len(row):
                continue
            v = str(row[j]).strip()
            if not v or v == "-":
                continue
            if _looks_like_percentage_value(v) and v not in seen_val:
                seen_val.add(v)
                ordered_pct.append((j, v))

        if len(ordered_pct) >= 2:
            vals = [v for _, v in ordered_pct]
            if (not pct1 or pct1 == "-") and (not pct2 or pct2 == "-"):
                row[idx_pct1] = vals[0]
                row[idx_pct2] = vals[1]
            elif not pct2 or pct2 == "-":
                # (1) sudah terisi, isi (2) dengan nilai lain
                other = next((v for v in vals if v != pct1), None)
                if other is not None:
                    row[idx_pct2] = other
            elif not pct1 or pct1 == "-":
                other = next((v for v in vals if v != pct2), None)
                if other is not None:
                    row[idx_pct1] = other
        elif len(ordered_pct) == 1 and (not pct2 or pct2 == "-"):
            # Satu persen ada; cek kolom Perubahan (kadang nilai (2) salah masuk sini)
            v17 = str(row[idx_perubahan]).strip() if idx_perubahan < len(row) else ""
            if _looks_like_percentage_value(v17):
                row[idx_pct2] = v17
                row[idx_perubahan] = "-"


def _fill_pct2_from_raw_lines(table: list[list[str]], raw_blue_lines: list[str]) -> None:
    """
    Isi Persentase (2) yang masih kosong dengan mencari di raw_blue_lines: setelah
    kemunculan nilai Persentase (1) baris ini, ambil token berikutnya yang mirip persen.
    Generik (tanpa hardcode No), untuk kasus nilai (2) tidak ikut ter-ekstrak ke tabel.
    """
    if not table or len(table) < 2 or not raw_blue_lines:
        return

    idx_pct1 = 13
    idx_pct2 = 16
    lines = [str(w).strip() for w in raw_blue_lines]
    search_start = 0
    look_ahead = 25

    for row in table[1:]:
        while len(row) <= idx_pct2:
            row.append("-")
        pct1 = str(row[idx_pct1]).strip() if idx_pct1 < len(row) else ""
        pct2 = str(row[idx_pct2]).strip() if idx_pct2 < len(row) else ""
        if not pct1 or pct1 == "-" or (pct2 and pct2 != "-"):
            continue

        # Cari pct1 di raw_blue_lines mulai dari search_start
        found_at = -1
        for i in range(search_start, len(lines)):
            if lines[i] == pct1:
                found_at = i
                break
        if found_at < 0:
            continue

        # Cari token mirip persen (bukan pct1) di look_ahead berikutnya
        for j in range(found_at + 1, min(found_at + 1 + look_ahead, len(lines))):
            tok = lines[j]
            if not tok or tok == "-":
                continue
            if _looks_like_percentage_value(tok) and tok != pct1:
                row[idx_pct2] = tok
                break

        search_start = found_at + 1


def _remove_spurious_merge_rows(table: list[list[str]]) -> list[list[str]]:
    """
    Hapus baris "hantu" hasil salah parse merge cell (generik, tidak hardcode No).
    Baris spurious: nilai dari baris merge (Persentase (1), Persentase (2)) masuk ke
    baris detail sehingga Perubahan berisi persen, Persentase (2) kosong,
    Jumlah Saham (2) atau Saham Gabungan (2) berisi persen. Baris seperti ini dihapus.
    Sebelum hapus: nilai yang salah tempat (persen di Perubahan / J2/SG2) dikembalikan
    ke baris sebelumnya (baris merge utama) agar kolom Persentase (1)/(2) dan
    Saham Gabungan (1)/(2) tidak kosong.
    """
    if not table or len(table) < 2:
        return table

    # Struktur 18 kolom
    idx_no = 0
    idx_j1 = 11
    idx_sg1 = 12
    idx_pct1 = 13
    idx_j2 = 14
    idx_sg2 = 15
    idx_pct2 = 16
    idx_perubahan = 17

    header = table[0]
    ncols = len(header)
    if ncols <= idx_perubahan:
        return table

    def ensure_len(row: list[str], length: int) -> list[str]:
        r = list(row)
        while len(r) < length:
            r.append("-")
        return r

    def is_spurious(row: list[str]) -> bool:
        row = ensure_len(row, ncols)
        pct2 = str(row[idx_pct2]).strip() if idx_pct2 < len(row) else ""
        perubahan = str(row[idx_perubahan]).strip() if idx_perubahan < len(row) else ""
        j2 = str(row[idx_j2]).strip() if idx_j2 < len(row) else ""
        sg2 = str(row[idx_sg2]).strip() if idx_sg2 < len(row) else ""
        if pct2 and pct2 != "-":
            return False
        if not _looks_like_percentage_value(perubahan):
            return False
        if _looks_like_percentage_value(j2) or _looks_like_percentage_value(sg2):
            return True
        return False

    def recover_to_prev(prev_row: list[str], curr_row: list[str]) -> None:
        """Isi baris sebelumnya (merge utama) dengan nilai yang salah tempat di baris spurious."""
        prev = ensure_len(prev_row, ncols)
        curr = ensure_len(curr_row, ncols)
        no_prev = str(prev[idx_no]).strip() if idx_no < len(prev) else ""
        no_curr = str(curr[idx_no]).strip() if idx_no < len(curr) else ""
        # Hanya jika baris spurious punya No sama atau kosong (lanjutan)
        if no_curr and no_curr != "-" and no_curr != no_prev:
            return
        # Persentase (2) dari kolom Perubahan baris spurious
        val_perubahan = str(curr[idx_perubahan]).strip() if idx_perubahan < len(curr) else ""
        if _looks_like_percentage_value(val_perubahan):
            prev[idx_pct2] = val_perubahan
        # Persentase (1) dari Jumlah Saham (2) atau Saham Gabungan (2) baris spurious
        val_j2 = str(curr[idx_j2]).strip() if idx_j2 < len(curr) else ""
        val_sg2 = str(curr[idx_sg2]).strip() if idx_sg2 < len(curr) else ""
        if _looks_like_percentage_value(val_j2):
            prev[idx_pct1] = val_j2
        elif _looks_like_percentage_value(val_sg2):
            prev[idx_pct1] = val_sg2
        # Saham Gabungan (1)/(2): jika baris spurious punya angka besar di J1/J2, bisa dipakai isi baris utama yang kosong
        for j_col, sg_col in ((idx_j1, idx_sg1), (idx_j2, idx_sg2)):
            val_j = str(curr[j_col]).strip() if j_col < len(curr) else ""
            prev_sg = str(prev[sg_col]).strip() if sg_col < len(prev) else ""
            if _looks_like_large_number(val_j) and (not prev_sg or prev_sg == "-"):
                prev[sg_col] = val_j
        # Salin kembali ke prev_row (mutate)
        for i in range(len(prev)):
            if i < len(prev_row):
                prev_row[i] = prev[i]
            else:
                prev_row.append(prev[i])
        while len(prev_row) > ncols:
            prev_row.pop()

    kept = [table[0]]
    i = 1
    while i < len(table):
        row = table[i]
        if is_spurious(row):
            # Sebelum hapus: kembalikan nilai ke baris sebelumnya (baris merge)
            if kept:
                prev_row = kept[-1]
                recover_to_prev(prev_row, row)
            i += 1
            continue
        kept.append(row)
        i += 1
    return kept


def _fill_saham_gabungan_dan_persentase_kosong(table: list[list[str]]) -> None:
    """
    Isi kolom Saham Gabungan (1)/(2) yang masih kosong dari Jumlah Saham (1)/(2)
    di baris yang sama (generik). Mutasi baris di table in-place.
    """
    if not table or len(table) < 2:
        return
    idx_sg1 = 12
    idx_sg2 = 15
    idx_j1 = 11
    idx_j2 = 14
    ncols = 18

    for row in table[1:]:
        while len(row) < ncols:
            row.append("-")
        # Saham Gabungan (1) kosong tapi Jumlah Saham (1) terisi angka besar -> salin
        sg1 = str(row[idx_sg1]).strip() if idx_sg1 < len(row) else ""
        j1 = str(row[idx_j1]).strip() if idx_j1 < len(row) else ""
        if (not sg1 or sg1 == "-") and _looks_like_large_number(j1):
            row[idx_sg1] = j1
        # Saham Gabungan (2) kosong tapi Jumlah Saham (2) terisi angka besar -> salin
        sg2 = str(row[idx_sg2]).strip() if idx_sg2 < len(row) else ""
        j2 = str(row[idx_j2]).strip() if idx_j2 < len(row) else ""
        if (not sg2 or sg2 == "-") and _looks_like_large_number(j2):
            row[idx_sg2] = j2


@app.route("/")
def index():
    """Halaman utama (re-use template yang sama)."""
    return render_template("index.html")


@app.route("/api/extract-blue", methods=["POST"])
def extract_blue():
    """
    Ekstrak teks biru, bangun tabel dari posisi, dan kembalikan JSON.

    Perbedaan utama dibanding `app.py`:
    - `table` dan `raw_data` langsung memakai `result["raw_data"]`
      dari `build_table_with_header_from_pdf`, sehingga:
        * struktur kolom selalu 18 kolom sesuai raw data,
        * isi tabel di front-end = isi raw data (lebih konsisten),
        * penanganan merge cell (Nama Emiten, Saham Gabungan Per Investor (1)/(2),
          Persentase Kepemilikan Per Investor (1)/(2), dll.) mengikuti logika
          di `build_table_with_header_from_pdf` yang sudah memperhatikan
          overlap beberapa baris (merge > 1–2 baris).
    """
    if "file" not in request.files:
        return jsonify({"error": "Tidak ada file"}), 400

    file = request.files["file"]
    if file.filename == "":
        return jsonify({"error": "File tidak dipilih"}), 400
    if not file.filename.lower().endswith(".pdf"):
        return jsonify({"error": "Hanya file PDF yang didukung"}), 400

    tmp_in_path: str | None = None
    try:
        # Simpan sementara PDF yang di-upload
        with tempfile.NamedTemporaryFile(delete=False, suffix=".pdf") as tmp_in:
            file.save(tmp_in.name)
            tmp_in_path = tmp_in.name

        # Bangun struktur tabel + raw_data memakai logika dari `app.py`
        result = build_table_with_header_from_pdf(tmp_in_path)
        if not result:
            return jsonify(
                {
                    "error": "Tidak ada teks warna biru ditemukan di PDF ini.",
                    "hint": "Pastikan teks benar-benar menggunakan warna biru (bukan hitam/abu-abu).",
                }
            ), 422

        # Raw teks biru: satu kata per baris (untuk debugging / PDF raw_teks_biru_satu_baris)
        raw_blue_lines: list[str] = []
        try:
            blue_spans = extract_blue_spans_with_bbox(tmp_in_path)
            for item in blue_spans:
                text = (item.get("text") or "").strip()
                for word in text.split():
                    if word:
                        raw_blue_lines.append(word)
        except Exception:
            raw_blue_lines = []

        # Jika build_table_with_header_from_pdf mengembalikan dict lengkap
        # (header_top, header_row, data, raw_data), gunakan raw_data sebagai sumber kebenaran.
        if isinstance(result, dict):
            raw_data = result.get("raw_data") or []
            header_top = result.get("header_top") or []

            # Jika raw_data kosong (fallback lama), bangun dari header_row + data
            if not raw_data:
                data_rows = result.get("data") or []
                header_row = result.get("header_row")
                if not header_row and data_rows:
                    # Tanpa header eksplisit, pakai panjang baris pertama
                    col_count = len(data_rows[0])
                    header_row = [f"Kolom {i+1}" for i in range(col_count)]
                header_row = list(header_row or [])
                col_count = len(header_row)
                table = [header_row]
                for row in data_rows:
                    r = (list(row) + ["-"] * col_count)[:col_count]
                    table.append(r)
                raw_data = table

            # Normalisasi: pastikan setiap baris punya panjang yang sama (berdasarkan baris pertama)
            max_cols = max(len(r) for r in raw_data) if raw_data else 0
            normalized_raw = []
            for r in raw_data:
                row = list(r)
                if max_cols:
                    row = (row + ["-"] * max_cols)[:max_cols]
                normalized_raw.append(row)

            # Hapus baris hantu DULU: baris detail yang salah dapat nilai merge (Persentase/Perubahan tertukar).
            # Harus sebelum pengisian persen agar baris spurious masih terdeteksi (Perubahan=persen, Pct2 kosong).
            # Nilai yang salah tempat dikembalikan ke baris merge sebelumnya.
            normalized_raw = _remove_spurious_merge_rows(normalized_raw)

            # Isi Persentase (1)/(2) yang kosong dari nilai persen lain di baris yang sama (generik).
            _fill_persentase_from_same_row(normalized_raw)

            # Jika Persentase (2) masih kosong, cari di raw teks biru (urutan kata) setelah nilai (1).
            _fill_pct2_from_raw_lines(normalized_raw, raw_blue_lines)

            # Isi Saham Gabungan (1)/(2) yang masih kosong dari Jumlah Saham (1)/(2) di baris yang sama.
            _fill_saham_gabungan_dan_persentase_kosong(normalized_raw)

            out = {
                # TABEL UTAMA: langsung dari raw_data (sudah termasuk header_18 + data_18)
                "table": normalized_raw,
                # header_top tetap dikirim agar bisa ditampilkan di atas tabel bila diperlukan
                "header_top": header_top,
                # raw_data juga dikirim sebagai referensi eksplisit
                "raw_data": normalized_raw,
                # raw teks biru satu kata per baris, untuk pembuatan raw_teks_biru_satu_baris.pdf
                "raw_blue_lines": raw_blue_lines,
            }
            return jsonify(out)

        # Jika build_table_with_header_from_pdf mengembalikan list biasa (fallback sederhana)
        table_rows = result or []
        # Normalisasi list-of-lists
        max_cols = max((len(r) for r in table_rows), default=0)
        normalized = []
        for r in table_rows:
            row = list(r)
            if max_cols:
                row = (row + ["-"] * max_cols)[:max_cols]
            normalized.append(row)
        return jsonify(
            {
                "table": normalized,
                "header_top": [],
                "raw_data": normalized,
                "raw_blue_lines": raw_blue_lines,
            }
        )

    except Exception as e:
        return jsonify({"error": str(e)}), 500
    finally:
        if tmp_in_path and os.path.exists(tmp_in_path):
            try:
                os.unlink(tmp_in_path)
            except Exception:
                pass


@app.route("/api/download-raw-blue-pdf", methods=["POST"])
def download_raw_blue_pdf():
    """Terima JSON { lines: ["kata1", "kata2", ...] } (raw teks biru, satu kata per baris), kembalikan PDF."""
    try:
        data = request.get_json() or {}
        lines = data.get("lines")
        if not lines or not isinstance(lines, list):
            return jsonify({"error": "Data lines tidak valid"}), 400
        lines = [str(x).strip() for x in lines if str(x).strip()]

        with tempfile.NamedTemporaryFile(suffix=".pdf", delete=False) as tmp_out:
            tmp_path = tmp_out.name

        create_pdf_raw_blue_one_per_line(lines, tmp_path)

        with open(tmp_path, "rb") as f:
            pdf_bytes = f.read()

        try:
            os.unlink(tmp_path)
        except Exception:
            pass

        return send_file(
            BytesIO(pdf_bytes),
            mimetype="application/pdf",
            as_attachment=True,
            download_name="raw_teks_biru_satu_baris.pdf",
        )
    except Exception as e:
        return jsonify({"error": str(e)}), 500


@app.route("/api/download-pdf", methods=["POST"])
def download_pdf():
    """Terima JSON { table: [[...], ...] }, kembalikan PDF berisi tabel."""
    try:
        data = request.get_json() or {}
        table = data.get("table")
        if not table or not isinstance(table, list):
            return jsonify({"error": "Data tabel tidak valid"}), 400

        rows = [r if isinstance(r, list) else [str(r)] for r in table]
        out_path = tempfile.mktemp(suffix=".pdf")
        try:
            create_pdf_from_table(rows, out_path)
            buf = BytesIO()
            with open(out_path, "rb") as f:
                buf.write(f.read())
            buf.seek(0)
            return send_file(
                buf,
                mimetype="application/pdf",
                as_attachment=True,
                download_name="teks_biru_tabel.pdf",
            )
        finally:
            if os.path.exists(out_path):
                try:
                    os.unlink(out_path)
                except Exception:
                    pass
    except Exception as e:
        return jsonify({"error": str(e)}), 500


if __name__ == "__main__":
    # Jalankan app2 pada port berbeda agar tidak bentrok dengan app.py (jika berjalan bersamaan).
    app.run(host="0.0.0.0", port=5001, debug=True)

