"""
Backend Flask: upload PDF, ekstrak teks warna biru, generate PDF output.
"""
import os
import re
import tempfile
from io import BytesIO
from flask import Flask, request, send_file, render_template, jsonify
import fitz  # PyMuPDF

app = Flask(__name__)
app.config["MAX_CONTENT_LENGTH"] = 16 * 1024 * 1024  # 16 MB max upload

# Warna biru: sRGB 0xRRGGBB atau tuple (r,g,b) 0–1. Kita anggap "biru" jika B dominan.
def _color_to_rgb(color):
    """Ubah color (int/tuple) ke (r, g, b) 0–255, atau None jika invalid."""
    if color is None:
        return None
    try:
        if isinstance(color, int) and not isinstance(color, bool):
            return ((color >> 16) & 0xFF, (color >> 8) & 0xFF, color & 0xFF)
        if isinstance(color, float):
            return None
        if hasattr(color, "__len__") and len(color) >= 3:
            return (int(float(color[0]) * 255), int(float(color[1]) * 255), int(float(color[2]) * 255))
    except (TypeError, ValueError, IndexError):
        pass
    return None


def is_blue_color(color) -> bool:
    """Cek apakah warna (sRGB int atau tuple float) dianggap biru."""
    rgb = _color_to_rgb(color)
    if rgb is None:
        return False
    r, g, b = rgb
    return b > r and b > g and b >= 80


def is_explicitly_other_color(color) -> bool:
    """True jika warna bukan biru — jangan ikutkan sebagai lanjutan paragraf biru.
    Hitam/abu-abu/merah/hijau = hentikan paragraf biru. None = tidak ada metadata (tetap hentikan agar teks hitam tidak ikut).
    """
    rgb = _color_to_rgb(color)
    if rgb is None:
        return True  # tidak ada warna = anggap bukan biru, jangan ikutkan (hindari teks hitam ikut)
    r, g, b = rgb
    if b > r and b > g and b >= 80:
        return False  # biru
    return True  # hitam, abu-abu, merah, hijau, dll = jangan ikutkan


def _as_list(val, default=None):
    """Pastikan nilai bisa di-iterate sebagai list."""
    if default is None:
        default = []
    if isinstance(val, list):
        return val
    if isinstance(val, (tuple, range)):
        return list(val)
    return default


def _span_to_item(span: dict, page_num: int) -> dict:
    raw_size = span.get("size", 12)
    try:
        size = float(raw_size) if raw_size is not None else 12
    except (TypeError, ValueError):
        size = 12
    bbox = span.get("bbox")
    if isinstance(bbox, (list, tuple)) and len(bbox) >= 4:
        x0, y0, x1, y1 = float(bbox[0]), float(bbox[1]), float(bbox[2]), float(bbox[3])
    else:
        x0 = y0 = x1 = y1 = 0
    return {
        "text": (span.get("text") or "").strip(),
        "size": size,
        "font": span.get("font", "helv"),
        "page": page_num + 1,
        "bbox": (x0, y0, x1, y1),
    }


def _flush_paragraph(current: list[dict], out: list[dict]) -> None:
    """Gabungkan semua span di current jadi satu paragraf, append ke out."""
    if not current:
        return
    lines = [it["text"] for it in current if it["text"]]
    if not lines:
        return
    merged = "\n".join(lines)
    out.append({
        "text": merged,
        "size": current[0]["size"],
        "font": current[0]["font"],
        "page": current[0]["page"],
    })


def extract_blue_text_from_pdf(input_path: str) -> list[dict]:
    """Baca PDF, kembalikan list paragraf biru. Satu paragraf = satu blok teks
    (banyak baris digabung). Span dalam blok yang sama digabung jadi satu item.
    Nomor halaman diambil dari halaman yang sedang diproses (page_num).
    """
    doc = fitz.open(input_path)
    blue_spans = []
    current_paragraph = []
    in_blue_paragraph = False
    for page_num in range(len(doc)):
        # Flush sisa paragraf dari halaman sebelumnya (jangan campur halaman)
        _flush_paragraph(current_paragraph, blue_spans)
        current_paragraph = []
        in_blue_paragraph = False

        page = doc[page_num]
        blocks = _as_list(page.get_text("dict", sort=True).get("blocks"))
        for block in blocks:
            if not isinstance(block, dict) or block.get("type") != 0:
                continue
            _flush_paragraph(current_paragraph, blue_spans)
            current_paragraph = []
            in_blue_paragraph = False
            for line in _as_list(block.get("lines")):
                if not isinstance(line, dict):
                    continue
                for span in _as_list(line.get("spans")):
                    if not isinstance(span, dict):
                        continue
                    color = span.get("color")
                    if is_blue_color(color):
                        in_blue_paragraph = True
                        current_paragraph.append(_span_to_item(span, page_num))
                    elif in_blue_paragraph and not is_explicitly_other_color(color):
                        current_paragraph.append(_span_to_item(span, page_num))
                    else:
                        _flush_paragraph(current_paragraph, blue_spans)
                        current_paragraph = []
                        in_blue_paragraph = False
            _flush_paragraph(current_paragraph, blue_spans)
            current_paragraph = []
    _flush_paragraph(current_paragraph, blue_spans)
    doc.close()
    return blue_spans


def extract_blue_spans_with_bbox(input_path: str) -> list[dict]:
    """Ekstrak semua span biru beserta bbox (x0,y0,x1,y1) dan page, tanpa menggabung paragraf."""
    doc = fitz.open(input_path)
    out = []
    for page_num in range(len(doc)):
        page = doc[page_num]
        blocks = _as_list(page.get_text("dict", sort=True).get("blocks"))
        for block in blocks:
            if not isinstance(block, dict) or block.get("type") != 0:
                continue
            for line in _as_list(block.get("lines")):
                if not isinstance(line, dict):
                    continue
                for span in _as_list(line.get("spans")):
                    if not isinstance(span, dict):
                        continue
                    if not is_blue_color(span.get("color")):
                        continue
                    item = _span_to_item(span, page_num)
                    if item["text"]:
                        out.append(item)
    doc.close()
    return out


def extract_all_spans_with_bbox(input_path: str) -> list[dict]:
    """Ekstrak SEMUA span (biru dan non-biru) beserta bbox dan page, untuk deteksi header tabel."""
    doc = fitz.open(input_path)
    out = []
    for page_num in range(len(doc)):
        page = doc[page_num]
        blocks = _as_list(page.get_text("dict", sort=True).get("blocks"))
        for block in blocks:
            if not isinstance(block, dict) or block.get("type") != 0:
                continue
            for line in _as_list(block.get("lines")):
                if not isinstance(line, dict):
                    continue
                for span in _as_list(line.get("spans")):
                    if not isinstance(span, dict):
                        continue
                    item = _span_to_item(span, page_num)
                    item["is_blue"] = is_blue_color(span.get("color"))
                    if item["text"]:
                        out.append(item)
    doc.close()
    return out


# Toleransi untuk mengelompokkan baris (pts) dan jarak minimal untuk kolom baru (pts).
# ROW_Y_TOLERANCE kecil (2) agar baris sub-header dan header utama tidak digabung.
# Jangan dibuat 1 agar baris lain tidak terpecah (berantakan); baris kembar (318 atas/bawah)
# dipecah lewat _split_rows_duplicate_numeric() jika satu sel berisi dua nilai.
ROW_Y_TOLERANCE = 2
COLUMN_X_GAP = 12

# Kata kunci untuk mendeteksi baris header tabel (case-insensitive, minimal 3 cocok).
# Sesuai bentuk lengkap: No. Urut, No. AODI, Nama Emiten, Nama Pemegang Saham,
# Nama Pemegang Rekening Efek, Alamat, Alamat (Lanjutan), Kebangsaan, Domisili,
# Status (Lokal/Asing), Kepemilikan Per ..., Jumlah Saham, Saham Gabungan, Persentase, Perubahan.
HEADER_KEYWORDS = (
    "no. urut",
    "no.urut",
    "no. aodi",
    "no.aodi",
    "kode efek",
    "nama emiten",
    "nama pemegang saham",
    "nama pemegang rekening efek",
    "nama rekening efek",
    "alamat",
    "alamat (lanjutan)",
    "alamat lanjutan",
    "kebangsaan",
    "domisili",
    "status (lokal/asing)",
    "status lokal",
    "kepemilikan per",
    "jumlah saham",
    "saham gabungan per investor",
    "persentase kepemilikan",
    "perubahan",
)
# Sub-kolom standar untuk blok "Kepemilikan Per [tanggal]"
KEPEMILIKAN_SUBCOLUMNS = (
    "Jumlah Saham",
    "Saham Gabungan Per Investor",
    "Persentase Kepemilikan Per Investor (%)",
)
# Template header baris kedua (nama kolom): 17 kolom fix sesuai spesifikasi.
# Urutan: No, Kode Efek, ... Status, lalu 2× (Jumlah Saham, Saham Gabungan, Persentase), Perubahan.
TEMPLATE_HEADER_FIXED = (
    "No",
    "Kode Efek",
    "Nama Emiten",
    "Nama Pemegang Rekening Efek",
    "Nama Pemegang Saham",
    "Nama Rekening Efek",
    "Alamat",
    "Alamat (Lanjutan)",
    "Kebangsaan",
    "Domisili",
    "Status (Lokal/Asing)",
)
# Template header 18 kolom (sama dengan raw data: No … Alamat, Alamat (Lanjutan), … Perubahan).
# Dua set numerik diberi label (1) dan (2) agar tidak membingungkan.
TEMPLATE_HEADER_18 = (
    "No",
    "Kode Efek",
    "Nama Emiten",
    "Nama Pemegang Rekening Efek",
    "Nama Pemegang Saham",
    "Nama Rekening Efek",
    "Alamat",
    "Alamat (Lanjutan)",
    "Kebangsaan",
    "Domisili",
    "Status (Lokal/Asing)",
    "Jumlah Saham (1)",
    "Saham Gabungan Per Investor (1)",
    "Persentase Kepemilikan Per Investor (%) (1)",
    "Jumlah Saham (2)",
    "Saham Gabungan Per Investor (2)",
    "Persentase Kepemilikan Per Investor (%) (2)",
    "Perubahan",
)
# Pola untuk pisah "No" dan "Kode Efek" dari sel pertama (e.g. "143 ATLA" -> No=143, Kode Efek=ATLA)
NO_KODE_EFEK_PATTERN = re.compile(r"^\s*(\d+)\s+(.*)$", re.DOTALL)
# Judul dokumen (satu baris panjang) bukan header tabel: header punya banyak kolom
MIN_HEADER_CELLS = 5
MIN_HEADER_KEYWORD_MATCHES = 3
# Baris header utama HARUS punya minimal satu kata kunci inti (agar baris sub-header
# seperti "Kepemilikan Per 28-JAN-2026" / "Jumlah Saham" saja tidak terpilih)
CORE_HEADER_KEYWORDS = (
    "no. urut",
    "no.urut",
    "no. aodi",
    "no.aodi",
    "kode efek",
    "nama emiten",
    "nama pemegang saham",
    "nama pemegang rekening efek",
)


def _row_text_lower(row_spans: list[dict]) -> str:
    """Gabung teks semua span di baris jadi satu string lowercase."""
    parts = [ (s.get("text") or "").strip() for s in row_spans ]
    return " ".join(parts).lower()


def _row_cell_count(row_spans: list[dict]) -> int:
    """Hitung jumlah 'kolom' di baris (span yang terpisah jarak > COLUMN_X_GAP)."""
    if not row_spans:
        return 0
    count = 1
    last_x1 = None
    for s in row_spans:
        bbox = s.get("bbox") or (0, 0, 0, 0)
        x0, x1 = bbox[0], bbox[2]
        if last_x1 is not None and (x0 - last_x1) > COLUMN_X_GAP:
            count += 1
        last_x1 = x1
    return count


def _row_looks_like_header(row_spans: list[dict]) -> bool:
    """
    True jika baris mirip header tabel LENGKAP (No, Kode Efek, Nama Emiten, Alamat, ...):
    - HARUS mengandung minimal 1 kata kunci inti (No/Kode Efek/Nama Emiten/Nama Pemegang ...),
    - minimal MIN_HEADER_KEYWORD_MATCHES dari daftar penuh,
    - minimal MIN_HEADER_CELLS kolom,
    - BUKAN baris yang hanya "Kepemilikan Per DD-MMM-YYYY" (sub-header tanggal saja).
    """
    text = _row_text_lower(row_spans)
    if not text or len(text) < 3:
        return False
    if not any(kw in text for kw in CORE_HEADER_KEYWORDS):
        return False
    # Tolak baris yang hanya sub-header tanggal (Kepemilikan Per 28-JAN-2026) tanpa kolom utama
    if "kepemilikan per" in text and not any(kw in text for kw in ("nama emiten", "kode efek", "nama pemegang", "no. urut", "no. aodi")):
        return False
    matches = sum(1 for kw in HEADER_KEYWORDS if kw in text)
    if matches < MIN_HEADER_KEYWORD_MATCHES:
        return False
    if _row_cell_count(row_spans) < MIN_HEADER_CELLS:
        return False
    return True


def _looks_like_stock_code(s: str) -> bool:
    """True jika teks mirip kode saham (singkat, huruf besar, alfanumerik)."""
    if not s or len(s) > 10:
        return False
    t = s.strip().upper()
    if len(t) < 2:
        return False
    if t.isdigit():
        return False
    return all(c.isalnum() or c in ".-" for c in t)


def _looks_like_no(s: str) -> bool:
    """True jika teks mirip nomor urut (angka saja, boleh dengan koma)."""
    if not s:
        return False
    t = s.strip().replace(",", "").replace(" ", "")
    return t.isdigit() and len(t) <= 6


def _looks_like_company_name(s: str) -> bool:
    """True jika teks mirip nama emiten (panjang, ada Tbk/PT)."""
    if not s or len(s) < 10:
        return False
    t = s.strip()
    return "Tbk" in t or ", PT" in t or "PT " in t or ("," in t and len(t) > 15)


def _split_rows_duplicate_numeric(raw_data_rows: list[tuple], num_cols: int) -> list[tuple]:
    """
    Jika satu baris punya sel di blok numerik (11-17) yang berisi dua nilai digabung (mis. "215,279,500 2,000,000"),
    pecah jadi dua baris: baris pertama dapat nilai pertama, baris kedua dapat nilai kedua.
    Memperbaiki kasus No 318 ketika dua baris PDF tergabung dalam satu cluster Y.
    """
    if num_cols < 18:
        return raw_data_rows
    idx_start = 11
    idx_end = min(18, num_cols)
    result = []
    for row_meta in raw_data_rows:
        cluster_y, cells, page = (row_meta[0], list(row_meta[1]) if len(row_meta) > 1 else [], row_meta[2] if len(row_meta) > 2 else row_meta[0])
        cells = (cells + [""] * num_cols)[:num_cols]
        parts_by_col = {}
        has_duplicate = False
        for j in range(idx_start, idx_end):
            if j >= len(cells):
                break
            val = (cells[j] or "").strip()
            if not val or val == "-" or " " not in val:
                continue
            tok = val.split()
            if len(tok) < 2:
                continue
            first = tok[0].strip()
            second = (" ".join(tok[1:])).strip()
            if not first or not second:
                continue
            is_num = (
                _looks_like_large_number(first) or _looks_like_percentage_value(first) or _looks_like_change_value(first)
            ) and (
                _looks_like_large_number(second) or _looks_like_percentage_value(second) or _looks_like_change_value(second)
            )
            if is_num:
                parts_by_col[j] = (first, second)
                has_duplicate = True
        if not has_duplicate or not parts_by_col:
            result.append(row_meta)
            continue
        row1 = list(cells)
        row2 = list(cells)
        for j in range(idx_start, idx_end):
            if j in parts_by_col:
                a, b = parts_by_col[j]
                row1[j] = a
                row2[j] = b
            else:
                row2[j] = "-"
        for j in range(0, idx_start):
            if j < len(row2):
                row2[j] = row1[j]
        result.append((cluster_y, row1, page))
        result.append((cluster_y, row2, page))
    return result


def _merge_split_kode_emiten_rows(raw_data_rows: list[tuple], num_cols: int) -> list[tuple]:
    """
    Gabungkan baris yang terpecah:
    - Pattern A: baris i kolom 1 = Nama Emiten (salah), kolom 2 kosong; baris berikut ada Kode Efek → pindah Nama Emiten ke kolom 2, isi kolom 1 dengan Kode Efek.
    - Pattern B: baris i kolom 1 = Kode Efek, kolom 2 kosong; baris berikut kolom 1 = Nama Emiten → isi kolom 2 dengan Nama Emiten dari baris berikut.
    Jangan gabung jika baris berikut punya data numerik sendiri (Jumlah Saham (1)/(2), Perubahan berbeda), mis. No 318 atas vs 318 bawah.
    """
    if num_cols < 3:
        return raw_data_rows
    idx_numeric_start = 11
    idx_numeric_end = min(18, num_cols)
    result = []
    i = 0
    while i < len(raw_data_rows):
        row_meta = raw_data_rows[i]
        cells = list((row_meta[1] if len(row_meta) > 1 else []) + [""] * num_cols)[:num_cols]
        col0 = (cells[0] if len(cells) > 0 else "").strip()
        col1 = (cells[1] or "").strip()
        col2 = (cells[2] if len(cells) > 2 else "").strip()

        merged = False
        # Cari baris berikut (hingga 3 baris ke depan) untuk pola terpecah
        for k in range(1, min(4, len(raw_data_rows) - i)):
            next_meta = raw_data_rows[i + k]
            next_cells = list((next_meta[1] if len(next_meta) > 1 else []) + [""] * num_cols)[:num_cols]
            next_col0 = (next_cells[0] if len(next_cells) > 0 else "").strip()
            next_col1 = (next_cells[1] if len(next_cells) > 1 else "").strip()
            same_entity = (not next_col0 or next_col0 == col0)

            # Jangan gabung jika baris berikut punya data numerik sendiri (mis. 318 bawah: 2,000,000 vs 318 atas)
            next_has_own_numeric = False
            for j in range(idx_numeric_start, idx_numeric_end):
                if j >= len(next_cells):
                    break
                nv = str(next_cells[j] or "").strip()
                rv = str(cells[j] if j < len(cells) else "").strip()
                if nv and nv != "-" and nv != rv:
                    next_has_own_numeric = True
                    break
            if next_has_own_numeric:
                continue

            # Pattern A: baris ini kolom 1 = nama emiten (salah), kolom 2 kosong; baris lain punya kode efek
            if (
                col1
                and not col2
                and _looks_like_company_name(col1)
                and next_col1
                and _looks_like_stock_code(next_col1)
                and same_entity
            ):
                cells[1] = next_col1
                cells[2] = col1
                for j in range(num_cols):
                    if j not in (1, 2) and (not cells[j] or not str(cells[j]).strip()) and j < len(next_cells):
                        v = next_cells[j]
                        if v and str(v).strip():
                            cells[j] = v
                result.append((row_meta[0], cells, row_meta[2] if len(row_meta) > 2 else row_meta[0]))
                i += 1 + k
                merged = True
                break

            # Pattern B: baris ini kolom 1 = kode efek, kolom 2 kosong; baris berikut kolom 1 = nama emiten
            if (
                col1
                and not col2
                and _looks_like_stock_code(col1)
                and next_col1
                and _looks_like_company_name(next_col1)
                and same_entity
            ):
                cells[2] = next_col1
                for j in range(num_cols):
                    if j != 2 and (not cells[j] or not str(cells[j]).strip()) and j < len(next_cells):
                        v = next_cells[j]
                        if v and str(v).strip():
                            cells[j] = v
                result.append((row_meta[0], cells, row_meta[2] if len(row_meta) > 2 else row_meta[0]))
                i += 1 + k
                merged = True
                break

        if not merged:
            result.append((row_meta[0], cells, row_meta[2] if len(row_meta) > 2 else row_meta[0]))
            i += 1
    return result


def _fix_no_kode_efek_cells(cells: list, num_cols: int) -> None:
    """
    Koreksi in-place: kolom No (0) harus berisi nomor urut, bukan kode efek.
    - Jika sel 0 berisi "247 BKSL" (gabungan), pecah jadi No=247, Kode Efek=BKSL.
    - Jika sel 0 = kode efek dan sel 1 = nomor → tukar.
    - Jika sel 0 = kode efek dan sel 1 bukan nomor → pindah sel 0 ke sel 1 (Kode Efek), No kosong.
    """
    if num_cols < 2:
        return
    col0 = (cells[0] if len(cells) > 0 else "").strip()
    col1 = (cells[1] if len(cells) > 1 else "").strip()
    if not col0:
        return
    m = NO_KODE_EFEK_PATTERN.match(col0)
    if m:
        no_part, rest = m.group(1).strip(), m.group(2).strip()
        if no_part and rest:
            cells[0] = no_part
            if len(cells) <= 1:
                cells.append(rest)
            else:
                if not col1 or col1 == "-":
                    cells[1] = rest
                elif rest not in col1:
                    cells[1] = rest + " " + col1
        return
    if not _looks_like_stock_code(col0):
        return
    if _looks_like_no(col1):
        cells[0], cells[1] = col1, col0
        return
    cells[0] = "-"
    if not col1 or col1 == "-":
        cells[1] = col0


def _fix_kode_emiten_cells(cells: list, num_cols: int) -> None:
    """
    Koreksi in-place: jika kolom 1 (Kode Efek) berisi nama emiten dan kolom 2 (Nama Emiten) kosong,
    pindahkan isi kolom 1 ke kolom 2 dan kosongkan kolom 1. Agar Nama Emiten tidak kosong.
    """
    if num_cols < 3:
        return
    col1 = (cells[1] if len(cells) > 1 else "").strip()
    col2 = (cells[2] if len(cells) > 2 else "").strip()
    if col1 and not col2 and _looks_like_company_name(col1):
        cells[1] = "-"
        if len(cells) <= 2:
            cells.append(col1)
        else:
            cells[2] = col1


def _looks_like_percentage_value(s: str) -> bool:
    """True jika nilai mirip persentase (desimal seperti 5.00, 11.70), bukan bilangan bulat seperti 343 atau 0."""
    if not s or s.strip() == "-":
        return False
    s = s.strip().replace(",", "")
    if not s:
        return False
    try:
        v = float(s)
        # Persentase biasanya 0–100 dengan desimal; bilangan bulat besar (343) atau 0 = Perubahan
        if v == 0:
            return False
        if "." in s and abs(v) < 1000:
            return True
        return False
    except ValueError:
        return False


def _looks_like_text_not_number(s: str) -> bool:
    """True jika nilai jelas teks (nama, alamat, negara): ada huruf dan bukan murni angka/desimal."""
    if not s or s.strip() == "-":
        return False
    t = s.strip()
    # Angka dengan koma/point saja = bukan teks
    cleaned = t.replace(",", "").replace(".", "").replace(" ", "")
    if cleaned.isdigit():
        return False
    # Nilai desimal murni (5.02, 11.76) = bukan teks
    if _looks_like_percentage_value(t):
        return False
    # Ada huruf dan panjang > 2 = teks (nama, alamat, dll)
    return any(c.isalpha() for c in t) and len(t) > 2


def _looks_like_large_number(s: str) -> bool:
    """True jika nilai mirip bilangan besar (jumlah saham): angka dengan/tanpa koma atau titik (pemisah ribuan)."""
    if not s or s.strip() == "-":
        return False
    t = s.strip().replace(",", "").replace(" ", "").replace(".", "")  # titik/koma pemisah ribuan
    if not t.isdigit():
        return False
    return len(t) >= 4  # minimal orde ribuan


def _looks_like_address_or_wrong_text(s: str) -> bool:
    """True jika nilai mirip alamat atau teks yang salah tempat (bukan angka untuk Jumlah Saham/Saham Gabungan)."""
    if not s or s.strip() == "-":
        return False
    if _looks_like_large_number(s) or _looks_like_percentage_value(s) or _looks_like_change_value(s):
        return False
    t = s.strip().upper()
    if _looks_like_securities_name(s) or _looks_like_person_name(s) or _looks_like_text_not_number(s):
        return True
    if any(k in t for k in ("JL ", "JLN ", "KAV", "FLOOR", "RT/RW", "UNIT ", "GD ", "MENARA")):
        return True
    return len(s) > 30 and "," in s


def _looks_like_change_value(s: str) -> bool:
    """True jika nilai mirip kolom Perubahan: angka kecil, 0, atau '-'."""
    if not s:
        return True
    t = s.strip()
    if t == "-":
        return True
    t = t.replace(",", "").replace(" ", "")
    if not t.isdigit():
        return False
    return int(t) >= 0 and len(t) <= 15  # angka wajar untuk perubahan


def _fix_split_percentage_cells(cells: list, num_cols: int) -> None:
    """
    Jika sel Persentase (1) atau (2) berisi dua nilai digabung (mis. "34.05\\n37.826.100.852"),
    pisahkan: nilai persen tetap di kolom persen, nilai besar pindah ke Saham Gabungan.
    """
    if num_cols < 18:
        return
    for idx_pct, idx_saham in ((13, 12), (16, 15)):  # Persentase(1)+Saham Gab(1), Persentase(2)+Saham Gab(2)
        val = (cells[idx_pct] if idx_pct < len(cells) else "").strip() or "-"
        if "\n" not in val:
            continue
        parts = [p.strip() for p in val.replace("\n", " ").split() if p.strip()]
        pct_part = None
        large_part = None
        for p in parts:
            if _looks_like_percentage_value(p):
                pct_part = p
            elif _looks_like_large_number(p):
                large_part = p
        if pct_part and large_part:
            while len(cells) <= max(idx_pct, idx_saham):
                cells.append("-")
            cells[idx_pct] = pct_part
            if (cells[idx_saham] or "").strip() in ("-", ""):
                cells[idx_saham] = large_part
        elif pct_part:
            while len(cells) <= idx_pct:
                cells.append("-")
            cells[idx_pct] = pct_part


def _fix_jumlah_saham_split_percentage(cells: list, num_cols: int) -> None:
    """
    Jika kolom Jumlah Saham (1) atau (2) berisi "persen angka_besar" (mis. "34.05 37,826,100,852"),
    pisahkan: persen → Persentase yang sesuai, angka besar tetap di Jumlah Saham.
    Memperbaiki kasus dimana nilai persen masih terikat dengan angka besar di kolom Jumlah Saham.
    """
    if num_cols < 18:
        return
    idx_jumlah1 = 11  # Jumlah Saham (1)
    idx_jumlah2 = 14  # Jumlah Saham (2)
    idx_pct1 = 13  # Persentase (1)
    idx_pct2 = 16  # Persentase (2)
    
    for idx_jumlah, idx_pct in ((idx_jumlah1, idx_pct1), (idx_jumlah2, idx_pct2)):
        val_jumlah = (cells[idx_jumlah] if idx_jumlah < len(cells) else "").strip() or "-"
        if val_jumlah == "-" or not val_jumlah:
            continue
        
        # Cek apakah berisi nilai persen dan angka besar
        parts = val_jumlah.split()
        if len(parts) < 2:
            continue
        
        pct_part = None
        large_part = None
        
        # Cek bagian pertama apakah persen
        first_part = parts[0].strip()
        if _looks_like_percentage_value(first_part):
            pct_part = first_part
            # Gabungkan sisa bagian sebagai angka besar
            rest_parts = " ".join(parts[1:]).strip()
            if _looks_like_large_number(rest_parts):
                large_part = rest_parts
        
        # Jika tidak ditemukan di bagian pertama, cek semua bagian
        if not pct_part:
            for p in parts:
                if _looks_like_percentage_value(p):
                    pct_part = p
                elif _looks_like_large_number(p):
                    if not large_part:
                        large_part = p
        
        # Jika ditemukan persen dan Persentase kosong, pindahkan persen
        if pct_part:
            val_pct = (cells[idx_pct] if idx_pct < len(cells) else "").strip() or "-"
            if val_pct == "-" or not val_pct:
                while len(cells) <= idx_pct:
                    cells.append("-")
                cells[idx_pct] = pct_part
                # Update Jumlah Saham dengan angka besar saja (jika ada)
                if large_part:
                    cells[idx_jumlah] = large_part
                else:
                    # Jika tidak ada angka besar, hapus persen dari Jumlah Saham
                    remaining = [p for p in parts if p != pct_part]
                    if remaining:
                        cells[idx_jumlah] = " ".join(remaining)
                    else:
                        cells[idx_jumlah] = "-"


def _fix_perubahan_split_percentage_then_number(cells: list, num_cols: int) -> None:
    """
    Jika kolom Perubahan berisi "X Y" dengan X = persen (36.67, 5.24, 21.95) dan Y = angka
    (2.703.857.638, 6,784,500, 0), pisah: X → Persentase (2), Y → Perubahan.
    Memperbaiki kasus No 341, 418, 497 dimana nilai persen salah masuk ke Perubahan.
    """
    if num_cols < 18:
        return
    idx_pct2 = 16
    idx_perubahan = 17
    val17 = (cells[idx_perubahan] if idx_perubahan < len(cells) else "").strip() or "-"
    if not val17 or val17 == "-":
        return
    parts = val17.split()
    if len(parts) < 2:
        return
    first, second = parts[0].strip(), " ".join(parts[1:]).strip()
    if not _looks_like_percentage_value(first):
        return
    # Second bisa angka besar (2.703.857.638), dengan koma (6,784,500), atau 0
    second_clean = second.replace(",", "").replace(".", "").replace(" ", "")
    if not second_clean or not second_clean.isdigit():
        return
    while len(cells) <= idx_perubahan:
        cells.append("-")
    cells[idx_pct2] = first
    cells[idx_perubahan] = second


def _looks_like_securities_name(s: str) -> bool:
    """True jika teks mirip nama rekening efek/securities (mengandung PT, SEKURITAS, ASSET, dll)."""
    if not s or s.strip() == "-":
        return False
    t = s.strip()
    if len(t) < 5:
        return False
    t_upper = t.upper()
    # Kata kunci yang menunjukkan nama perusahaan/securities
    keywords = ["SEKURITAS", "ASSET", "PT", "PT.", "LTD", "S/A", "INTERNATIONAL", "INDONESIA", 
                "MIRAE", "MANDIRI", "SINARMAS", "CGS", "AJAIB", "INDOVEST", "ABADIMUKTI"]
    # Jika mengandung kata kunci perusahaan DAN bukan angka/persen
    has_keyword = any(kw in t_upper for kw in keywords)
    if has_keyword and not _looks_like_percentage_value(t) and not _looks_like_large_number(t):
        return True
    # Atau jika mengandung "PT" atau "PT." di awal atau tengah
    if ("PT" in t_upper or "PT." in t_upper) and len(t) > 5:
        return not _looks_like_percentage_value(t) and not _looks_like_large_number(t)
    return False


def _looks_like_person_name(s: str) -> bool:
    """True jika teks mirip nama orang (huruf kapital, beberapa kata, tidak mengandung PT/SEKURITAS)."""
    if not s or s.strip() == "-":
        return False
    t = s.strip()
    # Nama orang biasanya tidak mengandung kata kunci perusahaan
    if _looks_like_securities_name(t):
        return False
    # Harus ada huruf, beberapa kata (spasi), dan tidak murni angka
    words = t.split()
    # Minimal 2 kata untuk nama lengkap (contoh: "ANDRIANSYAH PRAYITNO", "ADITYA ANTONIUS")
    if len(words) < 2:
        return False
    # Setidaknya 2 kata dan mengandung huruf, bukan angka/persen
    has_letters = any(c.isalpha() for c in t)
    is_not_number = not _looks_like_percentage_value(t) and not _looks_like_large_number(t) and not _looks_like_change_value(t)
    # Nama orang biasanya tidak terlalu panjang (maks ~50 karakter)
    reasonable_length = len(t) <= 50
    # Setiap kata harus mengandung huruf (bukan hanya angka/tanda baca)
    all_words_have_letters = all(any(c.isalpha() for c in word) for word in words)
    # Contoh: "ANDRIANSYAH PRAYITNO" - 2 kata, huruf semua, panjang wajar
    # Contoh: "ADITYA ANTONIUS" - 2 kata, huruf semua
    return has_letters and is_not_number and reasonable_length and len(words) >= 2 and all_words_have_letters


def _fix_numeric_block_by_content(cells: list, num_cols: int) -> None:
    """
    Koreksi blok kolom numerik (11-17): Persentase (1)/(2) harus berisi nilai persen,
    Perubahan harus angka kecil atau '-'. Jika ada teks (nama/alamat) di kolom tersebut,
    cari nilai yang sesuai di SELURUH BARIS lalu tukar. Jika tidak ditemukan, set ke "-".
    Indeks 18-kolom: 11=Jumlah(1), 12=Saham Gab(1), 13=Persentase(1), 14=Jumlah(2), 15=Saham Gab(2), 16=Persentase(2), 17=Perubahan.
    """
    if num_cols < 18:
        return
    idx_pct1 = 13
    idx_pct2 = 16
    idx_perubahan = 17
    block_start, block_end = 11, 18  # 11..17

    def get(i: int) -> str:
        return (cells[i] if i < len(cells) else "").strip() or "-"

    # Perubahan berisi "36.67 2.703.857.638" atau "5.24 6,784,500" → pisah, persen ke Persentase (2)
    _fix_perubahan_split_percentage_then_number(cells, num_cols)
    # Pisah sel yang berisi "34.05\n37.826.100.852" dll
    _fix_split_percentage_cells(cells, num_cols)
    # Pisah nilai persen dari kolom Jumlah Saham yang berisi "persen angka_besar"
    _fix_jumlah_saham_split_percentage(cells, num_cols)

    # Jumlah Saham dan Saham Gabungan (11, 12, 14, 15) tidak boleh berisi teks (alamat, nama rekening efek)
    for idx_jar in (11, 12, 14, 15):
        if idx_jar >= num_cols:
            break
        val_jar = get(idx_jar)
        if val_jar != "-" and _looks_like_address_or_wrong_text(val_jar):
            while len(cells) <= idx_jar:
                cells.append("-")
            cells[idx_jar] = "-"

    # Jika Persentase (1) berisi teks (nama rekening efek, nama pemegang saham dll), cari nilai persen di SELURUH BARIS
    val13 = get(idx_pct1)
    # Deteksi lebih agresif: nama orang, nama securities, atau teks umum yang bukan angka/persen
    is_text_in_pct1 = (_looks_like_text_not_number(val13) or _looks_like_securities_name(val13) or 
                       _looks_like_person_name(val13) or
                       (val13 != "-" and not _looks_like_percentage_value(val13) and not _looks_like_large_number(val13) and len(val13) > 3))
    if is_text_in_pct1:
        swapped = False
        # Cari di seluruh baris (0-17), bukan hanya blok 11-17
        for j in range(num_cols):
            if j == idx_pct1 or j == idx_pct2:  # Skip kolom persen lainnya
                continue
            if _looks_like_percentage_value(get(j)):
                cells[idx_pct1], cells[j] = get(j), val13
                swapped = True
                break
        # Jika tidak ada nilai persen yang ditemukan, set ke "-" HANYA jika jelas teks (nama/securities)
        # Jangan set ke "-" jika mungkin angka besar yang salah tempat
        if not swapped and (_looks_like_securities_name(val13) or _looks_like_person_name(val13)):
            while len(cells) <= idx_pct1:
                cells.append("-")
            cells[idx_pct1] = "-"
    
    # Jika Persentase (1) kosong ("-") atau berisi angka besar yang salah, cari nilai persen di seluruh baris
    val13_after = get(idx_pct1)
    if val13_after == "-" or (not _looks_like_percentage_value(val13_after) and _looks_like_large_number(val13_after)):
        # Cari nilai persen di seluruh baris (termasuk Perubahan)
        for j in range(num_cols):
            if j == idx_pct1 or j == idx_pct2:  # Skip kolom persen lainnya
                continue
            val_j = get(j)
            if _looks_like_percentage_value(val_j):
                # Jika Persentase (1) kosong, pindahkan persen ke sana
                if val13_after == "-":
                    cells[idx_pct1] = val_j
                    if j == idx_perubahan:  # Jika dari Perubahan, kosongkan Perubahan
                        cells[idx_perubahan] = "-"
                    else:
                        cells[j] = "-"
                # Jika Persentase (1) berisi angka besar, tukar dengan persen
                elif _looks_like_large_number(val13_after):
                    cells[idx_pct1], cells[j] = val_j, val13_after
                break

    # KOREKSI KHUSUS: Hanya jika Persentase (1) berisi persen dan Persentase (2) kosong,
    # tapi kolom periode 2 terisi, DAN tidak ada nilai persen lain di baris,
    # DAN periode 1 TIDAK punya data (untuk memastikan nilai tersebut memang untuk periode 2)
    # CATATAN: Logika ini dibuat lebih konservatif untuk menghindari memindahkan nilai yang seharusnya tetap di (1)
    val13_check = get(idx_pct1)
    val16_check = get(idx_pct2)
    idx_jumlah1 = 11  # Jumlah Saham (1)
    idx_saham_gab1 = 12  # Saham Gabungan Per Investor (1)
    idx_jumlah2 = 14  # Jumlah Saham (2)
    idx_saham_gab2 = 15  # Saham Gabungan Per Investor (2)
    val_jumlah1 = get(idx_jumlah1)
    val_saham_gab1 = get(idx_saham_gab1)
    val_jumlah2 = get(idx_jumlah2)
    val_saham_gab2 = get(idx_saham_gab2)
    
    # Cek apakah ada nilai persen lain di baris (selain di Persentase (1))
    has_other_percentage = False
    for j in range(num_cols):
        if j == idx_pct1 or j == idx_pct2:
            continue
        if _looks_like_percentage_value(get(j)):
            has_other_percentage = True
            break
    
    # Cek apakah periode 1 punya data
    has_period1_data = ((val_jumlah1 != "-" and val_jumlah1 and _looks_like_large_number(val_jumlah1)) or
                        (val_saham_gab1 != "-" and val_saham_gab1 and _looks_like_large_number(val_saham_gab1)))
    
    # Jika Persentase (1) berisi persen, Persentase (2) kosong, periode 2 punya data,
    # periode 1 TIDAK punya data, DAN TIDAK ada nilai persen lain di baris, baru pindahkan
    has_period2_data = ((val_jumlah2 != "-" and val_jumlah2 and _looks_like_large_number(val_jumlah2)) or
                        (val_saham_gab2 != "-" and val_saham_gab2 and _looks_like_large_number(val_saham_gab2)))
    
    if (_looks_like_percentage_value(val13_check) and 
        (val16_check == "-" or not val16_check) and
        has_period2_data and
        not has_period1_data and  # TAMBAHAN: periode 1 tidak punya data
        not has_other_percentage):  # Hanya jika TIDAK ada nilai persen lain
        # Pindahkan nilai dari Persentase (1) ke Persentase (2)
        while len(cells) <= idx_pct2:
            cells.append("-")
        cells[idx_pct2] = val13_check
        cells[idx_pct1] = "-"
    
    # KOREKSI: Jika ada nilai persen di kolom yang salah, pindahkan ke kolom Persentase yang sesuai.
    # Urutan penempatan Persentase (1)/(2) sudah mengikuti kiri-ke-kanan di tahap ekstraksi; di sini hanya perbaiki salah kolom.
    val13_final = get(idx_pct1)
    val16_final = get(idx_pct2)
    percentages_found = []
    
    # Kumpulkan semua nilai persen di baris (termasuk yang sudah di kolom Persentase) — format (idx, str, in_col)
    if _looks_like_percentage_value(val13_final):
        percentages_found.append((idx_pct1, val13_final, True))
    if _looks_like_percentage_value(val16_final):
        percentages_found.append((idx_pct2, val16_final, True))
    for j in range(num_cols):
        if j == idx_pct1 or j == idx_pct2:
            continue
        val_j = get(j)
        if _looks_like_percentage_value(val_j):
            percentages_found.append((j, val_j, False))
    
    # Jika Persentase (1) kosong dan ada nilai persen di kolom lain, pindahkan yang pertama ditemukan
    # Jangan pindahkan dari Persentase (2) jika periode 2 punya data
    if val13_final == "-" and len(percentages_found) > 0:
        # Cek apakah periode 2 punya data
        val_jumlah2 = get(idx_jumlah2)
        val_saham_gab2 = get(idx_saham_gab2)
        has_period2_data = ((val_jumlah2 != "-" and val_jumlah2 and _looks_like_large_number(val_jumlah2)) or
                            (val_saham_gab2 != "-" and val_saham_gab2 and _looks_like_large_number(val_saham_gab2)))
        
        # Cari nilai persen di kolom lain (bukan Persentase (2) jika periode 2 punya data)
        for pct_idx, pct_str, pct_in_col in percentages_found:
            # Jangan ambil dari Persentase (2) jika periode 2 punya data
            if pct_in_col and pct_idx == idx_pct2 and has_period2_data:
                continue
            # Jika ada nilai persen di kolom lain (bukan Persentase), pindahkan ke (1)
            if not pct_in_col and pct_idx != idx_pct2:
                cells[idx_pct1] = pct_str
                if pct_idx != idx_perubahan:
                    cells[pct_idx] = "-"
                break
            # Jika tidak ada nilai persen lain dan Persentase (2) tidak punya data periode 2, ambil dari (2)
            elif pct_in_col and pct_idx == idx_pct2 and not has_period2_data:
                cells[idx_pct1] = pct_str
                cells[idx_pct2] = "-"
                break
    
    # Jika Persentase (2) kosong dan ada nilai persen lain di kolom lain, pindahkan
    # TAPI: Prioritas jika periode 2 punya data, maka Persentase (2) harus terisi
    val16_after = get(idx_pct2)
    val_jumlah2_after = get(idx_jumlah2)
    val_saham_gab2_after = get(idx_saham_gab2)
    has_period2_data_after = ((val_jumlah2_after != "-" and val_jumlah2_after and _looks_like_large_number(val_jumlah2_after)) or
                              (val_saham_gab2_after != "-" and val_saham_gab2_after and _looks_like_large_number(val_saham_gab2_after)))
    
    if (val16_after == "-" or not val16_after):
        # Jika periode 2 punya data, cari nilai persen untuk Persentase (2)
        if has_period2_data_after:
            for j in range(num_cols):
                if j == idx_pct2 or j == idx_pct1:
                    continue
                val_j = get(j)
                if _looks_like_percentage_value(val_j):
                    while len(cells) <= idx_pct2:
                        cells.append("-")
                    cells[idx_pct2] = val_j
                    if j != idx_perubahan:
                        cells[j] = "-"
                    break
        # Jika periode 2 tidak punya data, gunakan logika lama (nilai persen kedua)
        elif len(percentages_found) > 1:
            found_count = 0
            for pct_idx, pct_str, pct_in_col in percentages_found:
                if pct_in_col and pct_idx == idx_pct1:
                    found_count += 1
                    continue
                if not pct_in_col or pct_idx != idx_pct1:
                    found_count += 1
                    if found_count == 2:  # Ambil nilai persen kedua
                        while len(cells) <= idx_pct2:
                            cells.append("-")
                        cells[idx_pct2] = pct_str
                        if not pct_in_col and pct_idx != idx_perubahan:
                            cells[pct_idx] = "-"
                        break
    
    # PENGECEKAN FINAL: Jika Persentase (1) masih kosong setelah semua koreksi,
    # cari lagi nilai persen di seluruh baris (mungkin terlewat sebelumnya)
    # TAPI: Jangan pindahkan dari Persentase (2) jika periode 2 punya data
    val13_final_check = get(idx_pct1)
    if val13_final_check == "-":
        # Cek apakah periode 2 punya data
        val_jumlah2_check = get(idx_jumlah2)
        val_saham_gab2_check = get(idx_saham_gab2)
        has_period2_data_check = ((val_jumlah2_check != "-" and val_jumlah2_check and _looks_like_large_number(val_jumlah2_check)) or
                                  (val_saham_gab2_check != "-" and val_saham_gab2_check and _looks_like_large_number(val_saham_gab2_check)))
        
        # Cari nilai persen di seluruh baris (termasuk yang mungkin terlewat)
        for j in range(num_cols):
            if j == idx_pct1:
                continue
            # Jangan ambil dari Persentase (2) jika periode 2 punya data
            if j == idx_pct2 and has_period2_data_check:
                continue
            val_j = get(j)
            if _looks_like_percentage_value(val_j):
                # Ambil nilai persen pertama yang ditemukan untuk Persentase (1)
                cells[idx_pct1] = val_j
                # Jangan kosongkan kolom sumber jika itu Perubahan atau Persentase (2) dengan data periode 2
                if j != idx_perubahan and not (j == idx_pct2 and has_period2_data_check):
                    cells[j] = "-"
                break
        
        # PENGECEKAN FINAL: Jika Persentase (2) kosong tapi periode 2 punya data, cari nilai persen untuk (2)
        val16_final_check = get(idx_pct2)
        val_jumlah2_final = get(idx_jumlah2)
        val_saham_gab2_final = get(idx_saham_gab2)
        has_period2_data_final = ((val_jumlah2_final != "-" and val_jumlah2_final and _looks_like_large_number(val_jumlah2_final)) or
                                  (val_saham_gab2_final != "-" and val_saham_gab2_final and _looks_like_large_number(val_saham_gab2_final)))
        
        if (val16_final_check == "-" or not val16_final_check) and has_period2_data_final:
            # Cari nilai persen di seluruh baris untuk Persentase (2)
            for j in range(num_cols):
                if j == idx_pct2 or j == idx_pct1:
                    continue
                val_j = get(j)
                if _looks_like_percentage_value(val_j):
                    while len(cells) <= idx_pct2:
                        cells.append("-")
                    cells[idx_pct2] = val_j
                    if j != idx_perubahan:
                        cells[j] = "-"
                    break

    # Jika Persentase (2) berisi teks, cari nilai persen di SELURUH BARIS
    val16 = get(idx_pct2)
    is_text_in_pct2 = (_looks_like_text_not_number(val16) or _looks_like_securities_name(val16) or 
                       (val16 != "-" and not _looks_like_percentage_value(val16) and not _looks_like_large_number(val16) and len(val16) > 3))
    if is_text_in_pct2:
        swapped = False
        for j in range(num_cols):
            if j == idx_pct1 or j == idx_pct2:  # Skip kolom persen lainnya
                continue
            if _looks_like_percentage_value(get(j)):
                cells[idx_pct2], cells[j] = get(j), val16
                swapped = True
                break
        # Jika tidak ada nilai persen yang ditemukan, set ke "-"
        if not swapped:
            while len(cells) <= idx_pct2:
                cells.append("-")
            cells[idx_pct2] = "-"

    # Jika Perubahan berisi teks (nama pemegang saham dll), lebih agresif
    val17 = get(idx_perubahan)
    is_text_in_perubahan = (_looks_like_text_not_number(val17) or _looks_like_person_name(val17) or 
                           _looks_like_securities_name(val17) or
                           (val17 != "-" and not _looks_like_percentage_value(val17) and 
                            not _looks_like_change_value(val17) and not _looks_like_large_number(val17) and len(val17) > 3))
    if is_text_in_perubahan:
        # Cari nilai angka yang cocok di blok 11-17 dulu
        swapped = False
        for j in range(block_start, block_end):
            if j == idx_perubahan:
                continue
            v = get(j)
            if _looks_like_change_value(v) and not _looks_like_large_number(v) and not _looks_like_percentage_value(v):
                cells[idx_perubahan], cells[j] = v, val17
                swapped = True
                break
        # Jika tidak ada atau jelas nama orang/securities, set ke "-"
        if not swapped or _looks_like_person_name(val17) or _looks_like_securities_name(val17):
            while len(cells) <= idx_perubahan:
                cells.append("-")
            cells[idx_perubahan] = "-"


def _fix_persentase_perubahan_cells(cells: list, num_cols: int) -> None:
    """
    Koreksi in-place: jika nilai di Perubahan (17) berbentuk persen (mis. 5.00, 11.70),
    pindahkan ke Persentase (1) atau (2) yang kosong. Prioritas: Persentase (1) dulu, lalu (2).
    Nilai bulat (343, 0) tetap di Perubahan.
    """
    if num_cols < 18:
        return
    idx_pct1 = 13
    idx_pct2 = 16
    idx_perubahan = 17
    val_perubahan = (cells[idx_perubahan] if len(cells) > idx_perubahan else "").strip()
    if not val_perubahan or val_perubahan == "-":
        return
    if not _looks_like_percentage_value(val_perubahan):
        return
    
    val_pct1 = (cells[idx_pct1] if len(cells) > idx_pct1 else "").strip() or "-"
    val_pct2 = (cells[idx_pct2] if len(cells) > idx_pct2 else "").strip() or "-"
    
    # Prioritas: pindahkan ke Persentase (1) jika kosong
    if val_pct1 == "-":
        while len(cells) <= idx_pct1:
            cells.append("-")
        cells[idx_pct1] = val_perubahan
        if len(cells) > idx_perubahan:
            cells[idx_perubahan] = "-"
    # Jika Persentase (1) sudah terisi, pindahkan ke Persentase (2) jika kosong
    elif val_pct2 == "-":
        while len(cells) <= idx_pct2:
            cells.append("-")
        cells[idx_pct2] = val_perubahan
        if len(cells) > idx_perubahan:
            cells[idx_perubahan] = "-"
    # Jika keduanya sudah terisi, biarkan di Perubahan (mungkin memang untuk Perubahan)


def _merge_continuation_rows(rows: list[list], num_cols: int) -> list[list]:
    """
    Gabungkan baris yang No-nya "-" (baris lanjutan/pecahan) ke baris sebelumnya,
    agar tidak jadi 2–3 baris terpisah. Isi baris lanjutan diisi ke kolom kosong
    baris utama, dari kanan (nilai terakhir isi Perubahan dll).
    """
    if num_cols < 2 or not rows:
        return rows
    result = []
    i = 0
    while i < len(rows):
        row = list(rows[i]) if rows[i] else []
        row = (row + ["-"] * num_cols)[:num_cols]
        while i + 1 < len(rows):
            next_row = list(rows[i + 1]) if rows[i + 1] else []
            next_row = (next_row + ["-"] * num_cols)[:num_cols]
            no_next = (next_row[0] or "").strip()
            if no_next and no_next != "-" and _looks_like_no(no_next):
                break
            # Jangan merge hanya jika baris lanjutan punya data numerik BEDA dari baris saat ini (baris data baru, mis. 497).
            # Jika numerik sama atau baris lanjutan hanya isi kosong → merge (lanjutan/duplikat, mis. 247/318 baris ke-3).
            idx_numeric_start = 11
            idx_numeric_end = min(18, num_cols)
            next_has_different_numeric = False
            for j in range(idx_numeric_start, idx_numeric_end):
                if j >= len(next_row):
                    break
                nv = str(next_row[j] or "").strip()
                rv = str(row[j] or "").strip()
                if not nv or nv == "-":
                    continue
                if not (
                    _looks_like_large_number(nv)
                    or _looks_like_percentage_value(nv)
                    or _looks_like_change_value(nv)
                ):
                    continue
                if rv and rv != "-" and nv != rv:
                    next_has_different_numeric = True
                    break
            if next_has_different_numeric:
                break
            empty_idx = [j for j in range(num_cols) if not (row[j] and str(row[j]).strip() and str(row[j]).strip() != "-")]
            values = [str(next_row[j]).strip() for j in range(num_cols) if next_row[j] and str(next_row[j]).strip() != "-"]
            if not values:
                i += 1
                continue
            if len(empty_idx) >= len(values):
                start = len(empty_idx) - len(values)
                for k, v in enumerate(values):
                    row[empty_idx[start + k]] = v
            else:
                for k, j in enumerate(empty_idx):
                    if k < len(values):
                        row[j] = values[k]
            i += 1
        result.append(row)
        i += 1
    return result


def _dedupe_rows_fill_kode_efek(rows: list[list], num_cols: int) -> list[list]:
    """
    Jika dua baris berurutan punya No sama, baris pertama Kode Efek kosong ("-") dan baris kedua
    punya Kode Efek: salin Kode Efek (dan kolom identitas lain yang kosong) ke baris pertama.
    Jangan gabung jadi satu baris jika baris kedua punya data numerik sendiri (Jumlah Saham (1)/(2),
    Perubahan berbeda), mis. No 318 atas vs 318 bawah — tetap pertahankan kedua baris.
    """
    if num_cols < 3 or not rows:
        return rows
    # Indeks kolom numerik (18-kolom): 11=Jumlah(1), 12=Saham Gab(1), 13=Persentase(1), 14=Jumlah(2), 15=Saham Gab(2), 16=Persentase(2), 17=Perubahan
    idx_numeric_start = 11
    idx_numeric_end = min(18, num_cols)
    result = []
    i = 0
    while i < len(rows):
        row = list(rows[i]) if rows[i] else []
        row = (row + ["-"] * num_cols)[:num_cols]
        if i + 1 < len(rows):
            next_row = list(rows[i + 1]) if rows[i + 1] else []
            next_row = (next_row + ["-"] * num_cols)[:num_cols]
            no_cur = (row[0] or "").strip()
            no_next = (next_row[0] or "").strip()
            kode_cur = (row[1] or "").strip()
            kode_next = (next_row[1] or "").strip()
            if no_cur and no_cur == no_next and (not kode_cur or kode_cur == "-") and kode_next and _looks_like_stock_code(kode_next):
                # Cek apakah baris kedua punya data numerik sendiri (bukan merge cell)
                second_has_own_numeric = False
                for j in range(idx_numeric_start, idx_numeric_end):
                    if j >= len(next_row):
                        break
                    nv = str(next_row[j] or "").strip()
                    rv = str(row[j] or "").strip()
                    if nv and nv != "-" and nv != rv:
                        second_has_own_numeric = True
                        break
                if second_has_own_numeric:
                    # Jangan gabung: isi hanya kolom identitas (0–10) di baris pertama yang kosong, pertahankan kedua baris
                    for j in range(1, min(11, num_cols)):
                        if (not row[j] or str(row[j]).strip() == "-") and j < len(next_row) and next_row[j] and str(next_row[j]).strip() != "-":
                            row[j] = next_row[j]
                    result.append(row)
                    result.append(next_row)
                    i += 2
                    continue
                # Baris kedua redundan: isi semua kolom kosong baris pertama dari baris kedua, buang baris kedua
                row[1] = kode_next
                for j in range(2, num_cols):
                    if (not row[j] or str(row[j]).strip() == "-") and j < len(next_row) and next_row[j] and str(next_row[j]).strip() != "-":
                        row[j] = next_row[j]
                i += 2
                result.append(row)
                continue
        result.append(row)
        i += 1
    return result


def _remove_duplicate_data_rows(raw_data_rows: list[tuple], num_cols: int) -> list[tuple]:
    """
    Hapus hanya baris yang benar-benar duplikat (seluruh isi baris sama).
    Jangan hapus hanya karena No + Kode Efek sama (satu No bisa punya dua baris: atas/bawah
    dengan Jumlah Saham dan Perubahan berbeda, mis. No 318).
    """
    if not raw_data_rows:
        return raw_data_rows
    result = []
    for row_meta in raw_data_rows:
        cells = list((row_meta[1] if len(row_meta) > 1 else []) + [""] * num_cols)[:num_cols]
        if result:
            prev_cells = list((result[-1][1] if len(result[-1]) > 1 else []) + [""] * num_cols)[:num_cols]
            if len(prev_cells) == len(cells) and all(
                (str(prev_cells[j] or "").strip() == str(cells[j] or "").strip() for j in range(len(cells)))
            ):
                continue
        result.append(row_meta)
    return result


def _group_spans_into_rows(span_items: list[dict]) -> list[tuple[float, int, list[dict]]]:
    """Kelompokkan span jadi baris: (y_mid, page, list span terurut x). Urutkan (page, y)."""
    if not span_items:
        return []
    def key(s):
        p = s.get("page", 1)
        b = s.get("bbox") or (0, 0, 0, 0)
        y = (b[1] + b[3]) / 2
        x = b[0]
        return (p, y, x)
    sorted_spans = sorted(span_items, key=key)
    rows = []
    current_y = None
    current_page = None
    current_row = []
    for s in sorted_spans:
        bbox = s.get("bbox") or (0, 0, 0, 0)
        mid_y = (bbox[1] + bbox[3]) / 2
        page = s.get("page", 1)
        if current_y is not None and (page != current_page or abs(mid_y - current_y) > ROW_Y_TOLERANCE):
            if current_row:
                rows.append((current_y, current_page, current_row))
            current_row = []
        current_y = mid_y
        current_page = page
        current_row.append(s)
    if current_row:
        rows.append((current_y, current_page, current_row))
    return rows


def build_table_with_header_from_pdf(input_path: str) -> list[list[str]]:
    """
    Pendekatan SEDERHANA dan LANGSUNG:
    1. Ambil semua teks biru dari PDF
    2. Kelompokkan per baris berdasarkan Y position
    3. Untuk setiap baris, ambil semua spans biru, urutkan berdasarkan X
    4. Tempatkan ke kolom berdasarkan posisi X relatif terhadap column boundaries dari header
    5. Deteksi merge cell dan duplicate ke semua baris yang ter-merge
    6. Setiap baris HARUS punya 18 kolom (tidak boleh kosong, paling hanya "-")
    """
    all_spans = extract_all_spans_with_bbox(input_path)
    if not all_spans:
        return []
    rows_raw = _group_spans_into_rows(all_spans)
    if not rows_raw:
        return []

    # Cari baris header
    header_row_idx = None
    for i, (_y, _page, row_spans) in enumerate(rows_raw):
        if _row_looks_like_header(row_spans):
            header_row_idx = i
            break
    if header_row_idx is None:
        blue_only = [s for s in all_spans if s.get("is_blue")]
        return build_table_from_spans(blue_only)

    header_spans = rows_raw[header_row_idx][2]

    # Kolom mengikuti jumlah alami dari header PDF (tanpa paksa merge/split) agar alignment benar
    TARGET_COLS = len(TEMPLATE_HEADER_18)  # 18 untuk output tampilan

    def _build_header_cells(gap: float) -> list[dict]:
        """Gabungkan span header menjadi cell berdasarkan gap X (satu cell per kolom nyata)."""
        spans_sorted = sorted(header_spans, key=lambda s: (s.get("bbox") or (0, 0, 0, 0))[0])
        cells: list[dict] = []
        cur_x0 = None
        cur_x1 = None
        cur_texts: list[str] = []
        for s in spans_sorted:
            bbox = s.get("bbox") or (0, 0, 0, 0)
            x0, x1 = float(bbox[0]), float(bbox[2])
            t = " ".join((s.get("text") or "").split())
            if not t:
                continue
            if cur_x1 is not None and x0 > (cur_x1 + gap):
                cells.append({"x0": cur_x0, "x1": cur_x1, "text": " ".join(cur_texts).strip()})
                cur_x0 = None
                cur_x1 = None
                cur_texts = []
            if cur_x0 is None:
                cur_x0 = x0
            cur_x1 = x1 if cur_x1 is None else max(cur_x1, x1)
            cur_texts.append(t)
        if cur_x1 is not None:
            cells.append({"x0": cur_x0, "x1": cur_x1, "text": " ".join(cur_texts).strip()})
        return [c for c in cells if (c.get("x1") is not None and c.get("x0") is not None and c.get("x1") > c.get("x0"))]

    # Pilih gap yang menghasilkan jumlah kolom terbanyak (jangan merge kolom yang terpisah di PDF)
    gap_candidates = [0.5, 1.0, 1.5, 2.0, 3.0, 4.0, 6.0, 8.0, 10.0, float(COLUMN_X_GAP)]
    best_cells = None
    for g in gap_candidates:
        c = _build_header_cells(g)
        if not c or len(c) < 8:
            continue
        if best_cells is None or len(c) > len(best_cells):
            best_cells = c
        if len(c) >= 20:
            break

    if not best_cells:
        blue_only = [s for s in all_spans if s.get("is_blue")]
        return build_table_from_spans(blue_only)

    header_cells = sorted(best_cells, key=lambda c: c["x0"])
    # ====== BEST PRACTICE: selalu buat 18 boundary kolom yang nyata ======
    # Banyak PDF menggabungkan header "Kepemilikan Per ..." sehingga header_cells < 18,
    # tapi data barisnya tetap punya 18 kolom. Jadi boundary tidak boleh cuma ikut header.
    #
    # Strategi:
    # - Jika header_cells >= 11 (kolom fixed sampai Status), gunakan boundary 11 kolom pertama dari header.
    # - Sisa lebar tabel (kanan setelah Status) dibagi rata menjadi 7 kolom: 3 + 3 + Perubahan.
    # - Jika header_cells >= 18, gunakan boundary 18 kolom dari header (paling presisi).
    # - Fallback: jika header_cells terlalu sedikit, bagi rata seluruh lebar tabel menjadi 18.

    blue_spans_all = [s for s in all_spans if s.get("is_blue") and s.get("bbox")]
    if blue_spans_all:
        data_x0_min = min(float((s.get("bbox") or (0, 0, 0, 0))[0]) for s in blue_spans_all)
        data_x1_max = max(float((s.get("bbox") or (0, 0, 0, 0))[2]) for s in blue_spans_all)
    else:
        data_x0_min = float(header_cells[0]["x0"])
        data_x1_max = float(header_cells[-1]["x1"])

    left_limit = min(float(header_cells[0]["x0"]), data_x0_min) - 2.0
    right_limit = max(float(header_cells[-1]["x1"]), data_x1_max) + 2.0

    edges: list[float] = []
    if len(header_cells) >= TARGET_COLS:
        # Boundary full dari header (18 kolom)
        hc = header_cells[:TARGET_COLS]
        edges.append(float(hc[0]["x0"]))
        for i in range(TARGET_COLS - 1):
            x1_current = float(hc[i]["x1"])
            x0_next = float(hc[i + 1]["x0"])
            edges.append((x1_current + x0_next) / 2)
        edges.append(float(hc[-1]["x1"]))
    elif len(header_cells) >= len(TEMPLATE_HEADER_FIXED):
        # Ambil 11 kolom fixed pertama dari header, lalu bagi kanan menjadi 7 kolom
        fixed_n = len(TEMPLATE_HEADER_FIXED)  # 11
        hc = header_cells[:fixed_n]
        edges.append(float(hc[0]["x0"]))
        for i in range(fixed_n - 1):
            x1_current = float(hc[i]["x1"])
            x0_next = float(hc[i + 1]["x0"])
            edges.append((x1_current + x0_next) / 2)
        status_right = float(hc[-1]["x1"])
        edges.append(status_right)
        # 7 kolom numerik/perubahan di kanan
        segs = TARGET_COLS - fixed_n  # 7
        span = max(1.0, right_limit - status_right)
        for i in range(1, segs + 1):
            edges.append(status_right + (i * span / segs))
    else:
        # Fallback: bagi rata seluruh lebar menjadi 18 kolom
        span = max(1.0, right_limit - left_limit)
        edges = [left_limit + (i * span / TARGET_COLS) for i in range(TARGET_COLS + 1)]

    # Normalisasi edges agar strictly increasing dan pakai left/right limit
    edges[0] = left_limit
    edges[-1] = right_limit
    for i in range(len(edges) - 1):
        if edges[i] >= edges[i + 1]:
            edges[i + 1] = edges[i] + 0.1

    column_boundaries = [(edges[i], edges[i + 1]) for i in range(TARGET_COLS)]
    num_cols = TARGET_COLS

    def _normalize_cell(s: str) -> str:
        """Trim dan satukan spasi/newline berlebih di isi sel."""
        if not s or not isinstance(s, str):
            return ""
        s = s.strip()
        # Jika hanya "-", kembalikan kosong
        if s == "-":
            return ""
        return " ".join(s.split())

    def column_index_for_span(bbox) -> int:
        """Tentukan kolom untuk span berdasarkan overlap area maksimum (posisi X di halaman)."""
        x0, _, x1, _ = bbox
        mid_x = (x0 + x1) / 2
        best_col = None
        max_overlap = 0.0
        
        for j, (cx0, cx1) in enumerate(column_boundaries):
            overlap_start = max(x0, cx0)
            overlap_end = min(x1, cx1)
            if overlap_start < overlap_end:
                overlap = overlap_end - overlap_start
                if overlap > max_overlap:
                    max_overlap = overlap
                    best_col = j
        
        if best_col is not None:
            return best_col
        # Fallback: kolom yang mengandung mid_x
        for j, (cx0, cx1) in enumerate(column_boundaries):
            if cx0 <= mid_x <= cx1:
                return j
            if mid_x < cx0:
                return max(0, j - 1)
        return num_cols - 1

    def _bbox_overlaps_col(bbox, col_idx: int) -> bool:
        x0, x1 = bbox[0], bbox[2]
        cx0, cx1 = column_boundaries[col_idx]
        return not (x1 <= cx0 or x0 >= cx1)

    # Header top dihapus - tidak perlu tampilkan "Kepemilikan Per tanggal"
    header_top = []

    # PENDEKATAN SEDERHANA: Ambil semua spans biru dari baris data, kelompokkan per baris
    # Kumpulkan semua spans biru dari baris data (setelah header)
    all_blue_spans = []
    for idx in range(header_row_idx + 1, len(rows_raw)):
        _y, _page, row_spans = rows_raw[idx]
        for s in row_spans:
            if s.get("is_blue"):
                bbox = s.get("bbox") or (0, 0, 0, 0)
                text = (s.get("text") or "").strip()
                if text and text != "-":
                    all_blue_spans.append({
                        "bbox": bbox,
                        "text": text,
                        "page": _page,
                        "y_mid": (bbox[1] + bbox[3]) / 2,
                        "x_mid": (bbox[0] + bbox[2]) / 2,
                        "y0": bbox[1],
                        "y1": bbox[3],
                    })
    
    if not all_blue_spans:
        blue_only = [s for s in all_spans if s.get("is_blue")]
        return build_table_from_spans(blue_only)
    
    # Hitung jarak baris normal untuk deteksi merge cell
    y_positions = sorted(set(s["y_mid"] for s in all_blue_spans))
    if len(y_positions) > 1:
        row_gaps = [y_positions[i+1] - y_positions[i] for i in range(len(y_positions)-1)]
        sorted_gaps = sorted(row_gaps)
        median_gap = sorted_gaps[len(sorted_gaps) // 2] if sorted_gaps else 10
        avg_row_gap = median_gap if median_gap > 0 else 10
    else:
        avg_row_gap = 10
    
    # Clustering Y positions: group Y yang berdekatan
    y_clusters = []
    for y_pos in y_positions:
        found = False
        for cluster_y in y_clusters:
            if abs(y_pos - cluster_y) <= ROW_Y_TOLERANCE:
                found = True
                break
        if not found:
            y_clusters.append(y_pos)
    y_clusters.sort()
    
    # Deteksi merge cells: span dengan tinggi lebih besar dari normal atau overlap dengan multiple clusters
    merged_cells_info = []
    for span in all_blue_spans:
        bbox = span["bbox"]
        y0, y1 = bbox[1], bbox[3]
        bbox_height = y1 - y0
        
        # Cari semua cluster Y yang overlap dengan bbox ini
        overlapping_clusters = []
        for cluster_y in y_clusters:
            if (y0 - ROW_Y_TOLERANCE <= cluster_y <= y1 + ROW_Y_TOLERANCE):
                overlapping_clusters.append(cluster_y)
        
        # Merge cell jika overlap dengan lebih dari 1 cluster Y atau tinggi > threshold
        is_merged = (len(overlapping_clusters) > 1 or bbox_height > avg_row_gap * 1.3 or bbox_height > 10)
        
        if is_merged:
            col_idx = column_index_for_span(bbox)
            if 0 <= col_idx < num_cols:
                merged_cells_info.append({
                    "col": col_idx,
                    "y0": y0,
                    "y1": y1,
                    "data": span["text"],
                    "page": span["page"],
                    "overlapping_clusters": overlapping_clusters
                })
    
    # Kelompokkan spans per baris berdasarkan cluster Y
    rows_by_cluster = {}  # {(page, cluster_y): [spans]}
    for span in all_blue_spans:
        page = span["page"]
        y_mid = span["y_mid"]
        
        # Cari cluster terdekat
        cluster_y = None
        min_dist = float('inf')
        for cy in y_clusters:
            dist = abs(y_mid - cy)
            if dist < min_dist:
                min_dist = dist
                cluster_y = cy
        
        if cluster_y is None:
            cluster_y = y_mid
        
        key = (page, cluster_y)
        if key not in rows_by_cluster:
            rows_by_cluster[key] = []
        rows_by_cluster[key].append(span)
    
    # BEST PRACTICE: Baca data teks biru kiri-ke-kanan, atas-ke-bawah.
    # Urutan baris = atas ke bawah (sorted_row_keys). Dalam tiap baris, urutan span = kiri ke kanan (sort by x_mid).
    # Untuk kolom Persentase (1) dan (2): nilai ditempatkan menurut urutan kemunculan (kiri ke kanan), bukan nilai terkecil.
    idx_pct1_col = 13  # Persentase Kepemilikan Per Investor (%) (1)
    idx_pct2_col = 16  # Persentase Kepemilikan Per Investor (%) (2)
    
    raw_data_rows = []
    sorted_row_keys = sorted(rows_by_cluster.keys(), key=lambda k: (k[0], k[1]))
    
    for (page, cluster_y) in sorted_row_keys:
        spans_in_row = rows_by_cluster[(page, cluster_y)]
        # Urutkan span kiri ke kanan (by x_mid) agar urutan baca = urutan di PDF
        spans_in_row = sorted(spans_in_row, key=lambda s: (s.get("x_mid") or 0))
        
        cells = [""] * num_cols
        # Kumpulkan semua nilai yang mirip persen di baris (dari kolom manapun) beserta posisi X
        pending_percentages = []  # list of (x_mid, text)
        
        for span in spans_in_row:
            text = (span.get("text") or "").strip()
            if not text:
                continue
            bbox = span.get("bbox") or (0, 0, 0, 0)
            x_mid = (bbox[0] + bbox[2]) / 2
            col_idx = column_index_for_span(bbox)
            if col_idx < 0:
                col_idx = 0
            if col_idx >= num_cols:
                col_idx = num_cols - 1
            
            # Setiap nilai yang mirip persen (di blok numerik 11-17) ditunda; ditempatkan menurut urutan kiri-kanan
            if col_idx >= 11 and _looks_like_percentage_value(text):
                pending_percentages.append((x_mid, text))
                continue
            
            # Penempatan biasa
            if cells[col_idx]:
                cells[col_idx] = cells[col_idx] + " " + text
            else:
                cells[col_idx] = text
        
        # Tempatkan nilai persen menurut urutan kiri ke kanan: pertama → Persentase (1), kedua → Persentase (2)
        pending_percentages.sort(key=lambda x: x[0])
        for i, (_, pct_text) in enumerate(pending_percentages):
            if i == 0:
                cells[idx_pct1_col] = pct_text
            elif i == 1:
                cells[idx_pct2_col] = pct_text
            # jika lebih dari 2, abaikan (kolom sudah terisi)
        
        cells = [_normalize_cell(c) for c in cells]
        if any(c.strip() for c in cells):
            raw_data_rows.append((cluster_y, cells, page))
    
    # Pecah baris yang punya dua nilai dalam satu sel numerik (mis. 318 atas/bawah tergabung)
    raw_data_rows = _split_rows_duplicate_numeric(raw_data_rows, num_cols)
    # Gabungkan baris terpecah: Kode Efek di baris bawah, Nama Emiten salah isi di kolom Kode Efek baris atas
    raw_data_rows = _merge_split_kode_emiten_rows(raw_data_rows, num_cols)
    # Hapus baris duplikat (No + Kode Efek sama)
    raw_data_rows = _remove_duplicate_data_rows(raw_data_rows, num_cols)
    
    # Duplicate merge cell data ke semua baris yang ter-merge
    if merged_cells_info and raw_data_rows:
        for merge_info in merged_cells_info:
            col_idx = merge_info["col"]
            merge_y0 = merge_info["y0"]
            merge_y1 = merge_info["y1"]
            merge_data = merge_info["data"]
            merge_page = merge_info["page"]
            overlapping_clusters = merge_info.get("overlapping_clusters", [])
            
            for row_idx, row_data in enumerate(raw_data_rows):
                row_cluster_y, row_cells, row_page = row_data
                
                if row_page != merge_page:
                    continue
                
                # Cek overlap
                is_overlapping = False
                if overlapping_clusters:
                    is_overlapping = row_cluster_y in overlapping_clusters
                else:
                    tolerance = avg_row_gap * 0.4
                    is_overlapping = (merge_y0 - tolerance <= row_cluster_y <= merge_y1 + tolerance)
                
                if is_overlapping:
                    while len(row_cells) <= col_idx:
                        row_cells.append("")
                    
                    current_cell_data = row_cells[col_idx].strip() if row_cells[col_idx] else ""
                    
                    # Kolom Jumlah Saham / Persentase / Perubahan (11-17) bukan merge: baris ke-2, ke-3, ... dalam blok merge
                    # jangan ditimpa agar nilai dari PDF (318 bawah, 247 bawah, dll.) tetap punya nilai sendiri.
                    if 11 <= col_idx <= 17:
                        is_second_or_later_in_block = False
                        if overlapping_clusters:
                            try:
                                pos_in_block = overlapping_clusters.index(row_cluster_y)
                                is_second_or_later_in_block = pos_in_block > 0
                            except ValueError:
                                pass
                        if not is_second_or_later_in_block and row_idx > 0:
                            prev_data = raw_data_rows[row_idx - 1]
                            prev_y = prev_data[0] if len(prev_data) > 0 else None
                            prev_page = prev_data[2] if len(prev_data) > 2 else row_page
                            prev_cells = prev_data[1] if len(prev_data) > 1 else []
                            if (prev_page == row_page and prev_y is not None and
                                merge_y0 - 1 <= prev_y <= merge_y1 + 1 and
                                merge_y0 - 1 <= row_cluster_y <= merge_y1 + 1 and
                                prev_y < row_cluster_y):
                                is_second_or_later_in_block = True
                            # Baris hasil split punya cluster_y sama; baris ke-2 jangan ditimpa (318 bawah: 2,000,000)
                            if prev_page == row_page and prev_y is not None and prev_y == row_cluster_y:
                                is_second_or_later_in_block = True
                            # Dua baris dengan No sama (318/707 atas-bawah): baris ke-2 jangan ditimpa
                            if prev_page == row_page and prev_cells and len(row_cells) > 0:
                                no_prev = (prev_cells[0] or "").strip()
                                no_cur = (row_cells[0] or "").strip()
                                if no_prev and no_prev == no_cur:
                                    is_second_or_later_in_block = True
                                # Baris lanjutan (No kosong) ikut baris sebelumnya: jangan timpa
                                if no_prev and (not no_cur or no_cur == "-"):
                                    is_second_or_later_in_block = True
                        if is_second_or_later_in_block:
                            raw_data_rows[row_idx] = (row_cluster_y, row_cells, row_page)
                            continue
                    
                    # Jangan timpa kolom numerik (11-17) jika sel sudah berisi angka yang wajar (mis. 2,000,000 di baris 318 bawah)
                    if 11 <= col_idx <= 17 and current_cell_data:
                        if (_looks_like_large_number(current_cell_data) or
                                _looks_like_percentage_value(current_cell_data) or
                                _looks_like_change_value(current_cell_data)):
                            raw_data_rows[row_idx] = (row_cluster_y, row_cells, row_page)
                            continue
                    
                    # Jumlah Saham (1)/(2) (kolom 11, 14): jangan timpa baris bawah (No sama atau kosong) dengan nilai baris atas
                    # agar 318 baris bawah tidak dapat 217,279,500/217,622,500 dari merge.
                    if col_idx in (11, 14) and row_idx > 0:
                        prev_cells = (raw_data_rows[row_idx - 1][1] if len(raw_data_rows[row_idx - 1]) > 1 else [])
                        no_prev = (prev_cells[0] or "").strip() if prev_cells else ""
                        no_cur = (row_cells[0] or "").strip() if row_cells else ""
                        if no_prev and (no_cur == no_prev or not no_cur or no_cur == "-"):
                            raw_data_rows[row_idx] = (row_cluster_y, row_cells, row_page)
                            continue
                    
                    # Jika kolom kosong atau merge_data lebih lengkap, gunakan merge_data
                    if not current_cell_data:
                        row_cells[col_idx] = merge_data
                    elif merge_data != current_cell_data and merge_data not in current_cell_data:
                        if len(merge_data) >= len(current_cell_data):
                            row_cells[col_idx] = merge_data
                    
                    raw_data_rows[row_idx] = (row_cluster_y, row_cells, row_page)
    
    # Konversi ke data_rows - langsung gunakan cells yang sudah ditempatkan ke kolom yang benar
    data_rows = []
    for row_data in raw_data_rows:
        if len(row_data) >= 2:
            _, row, _ = row_data if len(row_data) == 3 else (None, row_data[1] if len(row_data) > 1 else [], None)
            if row:
                # Pastikan row punya tepat num_cols kolom
                padded_row = (list(row) + [""] * num_cols)[:num_cols]
                data_rows.append(padded_row)

    # Finalisasi: pastikan selalu tepat 18 kolom, kosong = "-", dan koreksi Kode Efek/Nama Emiten
    template_header_row = list(TEMPLATE_HEADER_18)
    final_data_rows = []
    for row in data_rows:
        # Pad atau trim ke tepat 18 kolom
        padded_row = (list(row) + [""] * TARGET_COLS)[:TARGET_COLS]
        normalized_row = []
        for cell in padded_row:
            cell_str = str(cell).strip() if cell else ""
            if not cell_str or cell_str == "-":
                normalized_row.append("-")
            else:
                normalized_row.append(cell_str)
        # Jamin panjang tepat 18 (untuk baris yang mungkin kurang)
        while len(normalized_row) < TARGET_COLS:
            normalized_row.append("-")
        normalized_row = normalized_row[:TARGET_COLS]
        # Koreksi: kolom No jangan terisi kode efek (tukar/pecah/pindah)
        _fix_no_kode_efek_cells(normalized_row, TARGET_COLS)
        # Koreksi: jika Kode Efek berisi nama emiten dan Nama Emiten kosong, pindahkan
        _fix_kode_emiten_cells(normalized_row, TARGET_COLS)
        # Koreksi: data di kolom Perubahan (18) seharusnya di Persentase Kepemilikan % (17) hanya jika nilai mirip persen
        _fix_persentase_perubahan_cells(normalized_row, TARGET_COLS)
        # Koreksi: Persentase (1)/(2) dan Perubahan jangan berisi nama/alamat; tukar dengan nilai yang sesuai di blok 11-17
        _fix_numeric_block_by_content(normalized_row, TARGET_COLS)
        final_data_rows.append(normalized_row)

    # Gabungkan baris lanjutan (No "-") ke baris sebelumnya agar tidak jadi 2–3 baris terpisah
    final_data_rows = _merge_continuation_rows(final_data_rows, TARGET_COLS)
    # Rapikan Kode Efek: gabungkan baris duplikat No ketika baris pertama Kode Efek "-" dan baris kedua punya kode
    final_data_rows = _dedupe_rows_fill_kode_efek(final_data_rows, TARGET_COLS)
    
    # KOREKSI FINAL: Pastikan Persentase (1)/(2) dan Perubahan tidak berisi teks setelah merge/dedupe
    # (karena merge/dedupe bisa mengubah data, perlu koreksi ulang)
    for row in final_data_rows:
        _fix_numeric_block_by_content(row, TARGET_COLS)

    # BEST PRACTICE: Jangan gabung Alamat — tetap 18 kolom sesuai raw data (No … Perubahan).
    # Data teks biru sudah ditempatkan per kolom (column_boundaries); tiap baris = 18 sel berurutan.
    header_18 = list(template_header_row)
    data_18 = []
    for row in final_data_rows:
        row_18 = (list(row) + [""] * TARGET_COLS)[:TARGET_COLS]
        while len(row_18) < TARGET_COLS:
            row_18.append("-")
        row_18 = [str(c).strip() if c else "-" for c in row_18[:TARGET_COLS]]
        data_18.append(row_18)
    
    # Raw data untuk debugging: data mentah setelah ekstraksi (sebelum koreksi baris bawah dll.)
    raw_data_for_debug = [list(header_18)] + [list(row) for row in data_18]
    
    # Indeks 18 kolom (sesuai raw): 11=Jumlah(1), 12=Saham Gab(1), 13=Persentase(1), 14=Jumlah(2), 15=Saham Gab(2), 16=Persentase(2), 17=Perubahan
    TARGET_COLS_18 = TARGET_COLS  # 18
    idx_pct1_18 = 13
    idx_pct2_18 = 16
    idx_perubahan_18 = 17
    idx_jumlah1_18 = 11
    idx_saham_gab1_18 = 12
    idx_jumlah2_18 = 14
    idx_saham_gab2_18 = 15
    
    def get_18(i: int, cells: list) -> str:
        return (cells[i] if i < len(cells) else "").strip() or "-"
    
    def clear_cell_if_safe(r, col_idx):
        """Jangan kosongkan sel jika berisi angka besar (Jumlah Saham), agar nilai 2,000,000 tidak hilang."""
        if _looks_like_large_number(get_18(col_idx, r)):
            return
        r[col_idx] = "-"
    
    for idx_row, row_18 in enumerate(data_18):
        # Pastikan panjang 17 kolom
        while len(row_18) < TARGET_COLS_18:
            row_18.append("-")
        row_18 = row_18[:TARGET_COLS_18]
        data_18[idx_row] = row_18  # Pastikan perubahan tersimpan
        
        # Jumlah Saham dan Saham Gabungan (10, 11, 13, 14) jangan berisi alamat/nama rekening
        for j_jar in (idx_jumlah1_18, idx_saham_gab1_18, idx_jumlah2_18, idx_saham_gab2_18):
            v_jar = get_18(j_jar, row_18)
            if v_jar != "-" and _looks_like_large_number(v_jar) is False and _looks_like_address_or_wrong_text(v_jar):
                row_18[j_jar] = "-"
                data_18[idx_row] = row_18
        
        # Jika Persentase (1) berisi dua nilai (mis. "11.74 11.76"), pecah: (1)=nilai pertama, (2)=nilai kedua
        val12_split = get_18(idx_pct1_18, row_18)
        if val12_split and " " in val12_split:
            parts = val12_split.split()
            pcts = [p.strip() for p in parts if p.strip() and _looks_like_percentage_value(p.strip())]
            if len(pcts) >= 2 and (get_18(idx_pct2_18, row_18) == "-" or not get_18(idx_pct2_18, row_18)):
                row_18[idx_pct1_18] = pcts[0]
                row_18[idx_pct2_18] = pcts[1]
                data_18[idx_row] = row_18
        
        # Koreksi Persentase (1) - index 12
        val12 = get_18(idx_pct1_18, row_18)
        # Deteksi lebih agresif: jika bukan persen dan bukan angka besar, anggap teks
        # Termasuk nama orang ("ANDRIANSYAH PRAYITNO", "ADITYA ANTONIUS") dan nama securities
        is_not_pct = not _looks_like_percentage_value(val12)
        is_not_large_num = not _looks_like_large_number(val12)
        is_text_12 = (val12 != "-" and is_not_pct and is_not_large_num and 
                      (_looks_like_text_not_number(val12) or _looks_like_securities_name(val12) or 
                       _looks_like_person_name(val12) or len(val12) > 3))
        if is_text_12:
            swapped = False
            # Cari nilai persen di seluruh baris (0-16)
            for j in range(TARGET_COLS_18):
                if j == idx_pct1_18 or j == idx_pct2_18:
                    continue
                val_j = get_18(j, row_18)
                if _looks_like_percentage_value(val_j):
                    row_18[idx_pct1_18], row_18[j] = val_j, val12
                    swapped = True
                    break
            # Jika tidak ada nilai persen yang ditemukan, set ke "-" HANYA jika jelas teks (nama/securities)
            if not swapped and (_looks_like_securities_name(val12) or _looks_like_person_name(val12)):
                row_18[idx_pct1_18] = "-"
            data_18[idx_row] = row_18  # Simpan perubahan
        
        # Pisahkan nilai persen dari kolom Jumlah Saham (1) dan (2) jika ada
        for idx_jumlah_17, idx_pct_17 in ((idx_jumlah1_18, idx_pct1_18), (idx_jumlah2_18, idx_pct2_18)):
            val_jumlah_17 = get_18(idx_jumlah_17, row_18)
            if val_jumlah_17 != "-" and val_jumlah_17:
                parts = val_jumlah_17.split()
                if len(parts) >= 2:
                    first_part = parts[0].strip()
                    if _looks_like_percentage_value(first_part):
                        val_pct_17 = get_18(idx_pct_17, row_18)
                        if val_pct_17 == "-" or not val_pct_17:
                            row_18[idx_pct_17] = first_part
                            rest_parts = " ".join(parts[1:]).strip()
                            if _looks_like_large_number(rest_parts):
                                row_18[idx_jumlah_17] = rest_parts
                            else:
                                row_18[idx_jumlah_17] = rest_parts if rest_parts else "-"
                            data_18[idx_row] = row_18  # Simpan perubahan
        
        # Jika Persentase (1) kosong ("-") atau berisi angka besar yang salah, cari nilai persen di seluruh baris
        # TAPI: Jangan pindahkan dari Persentase (2) jika periode 2 punya data
        val12_after = get_18(idx_pct1_18, row_18)
        val_jumlah2_check = get_18(idx_jumlah2_18, row_18)
        val_saham_gab2_check = get_18(idx_saham_gab2_18, row_18)
        has_period2_data_check_17 = ((val_jumlah2_check != "-" and val_jumlah2_check and _looks_like_large_number(val_jumlah2_check)) or
                                     (val_saham_gab2_check != "-" and val_saham_gab2_check and _looks_like_large_number(val_saham_gab2_check)))
        
        if val12_after == "-" or (not _looks_like_percentage_value(val12_after) and _looks_like_large_number(val12_after)):
            # Cari nilai persen di seluruh baris (termasuk Perubahan), tapi skip Persentase (2) jika periode 2 punya data
            for j in range(TARGET_COLS_18):
                if j == idx_pct1_18 or j == idx_pct2_18:  # Skip kolom persen lainnya
                    continue
                # Jangan ambil dari Persentase (2) jika periode 2 punya data
                if j == idx_pct2_18 and has_period2_data_check_17:
                    continue
                val_j = get_18(j, row_18)
                if _looks_like_percentage_value(val_j):
                    # Jika Persentase (1) kosong, pindahkan persen ke sana
                    if val12_after == "-":
                        row_18[idx_pct1_18] = val_j
                        if j == idx_perubahan_18:  # Jika dari Perubahan, kosongkan Perubahan
                            row_18[idx_perubahan_18] = "-"
                        else:
                            clear_cell_if_safe(row_18, j)
                    # Jika Persentase (1) berisi angka besar, tukar dengan persen
                    elif _looks_like_large_number(val12_after):
                        row_18[idx_pct1_18], row_18[j] = val_j, val12_after
                    data_18[idx_row] = row_18  # Simpan perubahan
                    break
        
        # KOREKSI KHUSUS SETELAH MERGE ALAMAT: Hanya jika Persentase (1) berisi persen dan Persentase (2) kosong,
        # tapi kolom periode 2 terisi, periode 1 TIDAK punya data, DAN tidak ada nilai persen lain di baris
        val12_check = get_18(idx_pct1_18, row_18)
        val15_check = get_18(idx_pct2_18, row_18)
        val_jumlah1_17 = get_18(idx_jumlah1_18, row_18)
        val_saham_gab1_17 = get_18(idx_saham_gab1_18, row_18)
        val_jumlah2_17 = get_18(idx_jumlah2_18, row_18)
        val_saham_gab2_17 = get_18(idx_saham_gab2_18, row_18)
        
        # Cek apakah ada nilai persen lain di baris (selain di Persentase (1))
        has_other_percentage_17 = False
        for j in range(TARGET_COLS_18):
            if j == idx_pct1_18 or j == idx_pct2_18:
                continue
            if _looks_like_percentage_value(get_18(j, row_18)):
                has_other_percentage_17 = True
                break
        
        # Cek apakah periode 1 punya data
        has_period1_data_18 = ((val_jumlah1_17 != "-" and val_jumlah1_17 and _looks_like_large_number(val_jumlah1_17)) or
                                (val_saham_gab1_17 != "-" and val_saham_gab1_17 and _looks_like_large_number(val_saham_gab1_17)))
        
        has_period2_data_18 = ((val_jumlah2_17 != "-" and val_jumlah2_17 and _looks_like_large_number(val_jumlah2_17)) or
                                (val_saham_gab2_17 != "-" and val_saham_gab2_17 and _looks_like_large_number(val_saham_gab2_17)))
        
        if (_looks_like_percentage_value(val12_check) and 
            (val15_check == "-" or not val15_check) and
            has_period2_data_18 and
            not has_period1_data_18 and  # TAMBAHAN: periode 1 tidak punya data
            not has_other_percentage_17):  # Hanya jika TIDAK ada nilai persen lain
            # Pindahkan nilai dari Persentase (1) ke Persentase (2)
            row_18[idx_pct2_18] = val12_check
            row_18[idx_pct1_18] = "-"
            data_18[idx_row] = row_18  # Simpan perubahan
        
        # KOREKSI 318: Kedua periode punya data, (1) terisi tapi (2) kosong, dan ada nilai persen lain di baris.
        # Nilai yang sekarang di (1) seharusnya di (2); nilai di kolom lain (mis. Perubahan) seharusnya di (1).
        if (has_period1_data_18 and has_period2_data_18 and
            _looks_like_percentage_value(get_18(idx_pct1_18, row_18)) and
            (get_18(idx_pct2_18, row_18) == "-" or not get_18(idx_pct2_18, row_18))):
            pct_in_col1 = get_18(idx_pct1_18, row_18)
            other_pct_col = None
            other_pct_val = None
            for j in range(TARGET_COLS_18):
                if j == idx_pct1_18 or j == idx_pct2_18:
                    continue
                v = get_18(j, row_18)
                if _looks_like_percentage_value(v):
                    other_pct_col, other_pct_val = j, v
                    break
            if other_pct_col is not None and other_pct_val is not None:
                row_18[idx_pct1_18] = other_pct_val   # nilai dari kolom lain → (1)
                row_18[idx_pct2_18] = pct_in_col1     # nilai yang tadinya di (1) → (2)
                clear_cell_if_safe(row_18, other_pct_col)
                data_18[idx_row] = row_18
        
        # KOREKSI SETELAH MERGE ALAMAT: Jika ada nilai persen di kolom yang salah, pindahkan ke kolom Persentase yang sesuai.
        # Urutan Persentase (1)/(2) sudah mengikuti kiri-ke-kanan di tahap ekstraksi.
        val12_final = get_18(idx_pct1_18, row_18)
        val15_final = get_18(idx_pct2_18, row_18)
        percentages_found_17 = []
        if _looks_like_percentage_value(val12_final):
            percentages_found_17.append((idx_pct1_18, val12_final, True))
        if _looks_like_percentage_value(val15_final):
            percentages_found_17.append((idx_pct2_18, val15_final, True))
        for j in range(TARGET_COLS_18):
            if j == idx_pct1_18 or j == idx_pct2_18:
                continue
            val_j = get_18(j, row_18)
            if _looks_like_percentage_value(val_j):
                percentages_found_17.append((j, val_j, False))
        
        # Jika Persentase (1) kosong dan ada nilai persen di kolom lain, pindahkan yang pertama ditemukan
        if val12_final == "-" and len(percentages_found_17) > 0:
            # Cek apakah periode 2 punya data
            has_period2_data_for_pct1 = ((val_jumlah2_17 != "-" and val_jumlah2_17 and _looks_like_large_number(val_jumlah2_17)) or
                                        (val_saham_gab2_17 != "-" and val_saham_gab2_17 and _looks_like_large_number(val_saham_gab2_17)))
            
            # Cari nilai persen di kolom lain (bukan Persentase (2) jika periode 2 punya data)
            for pct_idx, pct_str, pct_in_col in percentages_found_17:
                # Jangan ambil dari Persentase (2) jika periode 2 punya data
                if pct_in_col and pct_idx == idx_pct2_18 and has_period2_data_for_pct1:
                    continue
                # Jika ada nilai persen di kolom lain (bukan Persentase), pindahkan ke (1)
                if not pct_in_col and pct_idx != idx_pct2_18:
                    row_18[idx_pct1_18] = pct_str
                    if pct_idx != idx_perubahan_18:
                        clear_cell_if_safe(row_18, pct_idx)
                    data_18[idx_row] = row_18  # Simpan perubahan
                    break
                # Jika tidak ada nilai persen lain dan Persentase (2) tidak punya data periode 2, ambil dari (2)
                elif pct_in_col and pct_idx == idx_pct2_18 and not has_period2_data_for_pct1:
                    row_18[idx_pct1_18] = pct_str
                    row_18[idx_pct2_18] = "-"
                    data_18[idx_row] = row_18  # Simpan perubahan
                    break
        
        # Jika Persentase (2) kosong dan ada nilai persen lain di kolom lain, pindahkan
        # Skip jika sudah ditangani oleh kasus khusus dua nilai persen dengan kedua periode punya data
        val15_after = get_18(idx_pct2_18, row_18)
        if (val15_after == "-" or not val15_after) and len(percentages_found_17) > 1:
            # Cari nilai persen kedua yang belum di kolom Persentase
            found_count = 0
            for pct_idx, pct_str, pct_in_col in percentages_found_17:
                if pct_in_col and pct_idx == idx_pct1_18:
                    found_count += 1
                    continue
                if not pct_in_col or pct_idx != idx_pct1_18:
                    found_count += 1
                    if found_count == 2:  # Ambil nilai persen kedua
                        row_18[idx_pct2_18] = pct_str
                        if not pct_in_col and pct_idx != idx_perubahan_18:
                            clear_cell_if_safe(row_18, pct_idx)
                        data_18[idx_row] = row_18  # Simpan perubahan
                        break
        
        # PENGECEKAN FINAL SETELAH MERGE ALAMAT: Jika Persentase (1) masih kosong setelah semua koreksi,
        # cari lagi nilai persen di seluruh baris (mungkin terlewat sebelumnya)
        # TAPI: Jangan isi Persentase (1) jika periode 1 TIDAK punya data (mis. baris 247: (1) harus tetap "-")
        val12_final_check = get_18(idx_pct1_18, row_18)
        if val12_final_check == "-" and has_period1_data_18:
            # Cek apakah periode 2 punya data
            has_period2_data_final = ((val_jumlah2_17 != "-" and val_jumlah2_17 and _looks_like_large_number(val_jumlah2_17)) or
                                     (val_saham_gab2_17 != "-" and val_saham_gab2_17 and _looks_like_large_number(val_saham_gab2_17)))
            
            # Cari nilai persen di seluruh baris (termasuk yang mungkin terlewat)
            for j in range(TARGET_COLS_18):
                if j == idx_pct1_18:
                    continue
                # Jangan ambil dari Persentase (2) jika periode 2 punya data
                if j == idx_pct2_18 and has_period2_data_final:
                    continue
                val_j = get_18(j, row_18)
                if _looks_like_percentage_value(val_j):
                    # Ambil nilai persen pertama yang ditemukan untuk Persentase (1)
                    row_18[idx_pct1_18] = val_j
                    # Jangan kosongkan kolom sumber jika itu Perubahan atau Persentase (2) dengan data periode 2, atau berisi angka besar
                    if j != idx_perubahan_18 and not (j == idx_pct2_18 and has_period2_data_final):
                        clear_cell_if_safe(row_18, j)
                    data_18[idx_row] = row_18  # Simpan perubahan
                    break
        # Paksa: jika periode 1 TIDAK punya data, Persentase (1) harus "-" (contoh: baris 247).
        # Jangan kosongkan jika baris sudah punya kedua persen (1) dan (2) dari ekstraksi PDF (mis. 318 atas).
        val1_now = get_18(idx_pct1_18, row_18)
        val2_now = get_18(idx_pct2_18, row_18)
        both_pct_filled = _looks_like_percentage_value(val1_now) and _looks_like_percentage_value(val2_now)
        if not has_period1_data_18 and not both_pct_filled:
            row_18[idx_pct1_18] = "-"
            data_18[idx_row] = row_18
        
        # PENGECEKAN: Jika Persentase (2) kosong tapi periode 2 punya data, cari nilai persen untuk (2)
        val15_final_check = get_18(idx_pct2_18, row_18)
        if (val15_final_check == "-" or not val15_final_check) and has_period2_data_18:
            # Cari nilai persen di seluruh baris untuk Persentase (2)
            for j in range(TARGET_COLS_18):
                if j == idx_pct2_18 or j == idx_pct1_18:
                    continue
                val_j = get_18(j, row_18)
                if _looks_like_percentage_value(val_j):
                    row_18[idx_pct2_18] = val_j
                    if j != idx_perubahan_18:
                        clear_cell_if_safe(row_18, j)
                    data_18[idx_row] = row_18  # Simpan perubahan
                    break
        
        # Koreksi Persentase (2) - index 15
        val15 = get_18(idx_pct2_18, row_18)
        # Deteksi lebih agresif: jika bukan persen dan bukan angka besar, anggap teks
        is_not_pct_15 = not _looks_like_percentage_value(val15)
        is_not_large_num_15 = not _looks_like_large_number(val15)
        is_text_15 = (val15 != "-" and is_not_pct_15 and is_not_large_num_15 and 
                      (_looks_like_text_not_number(val15) or _looks_like_securities_name(val15) or 
                       _looks_like_person_name(val15) or len(val15) > 3))
        if is_text_15:
            swapped = False
            # Cari nilai persen di seluruh baris (0-16)
            for j in range(TARGET_COLS_18):
                if j == idx_pct1_18 or j == idx_pct2_18:
                    continue
                val_j = get_18(j, row_18)
                if _looks_like_percentage_value(val_j):
                    row_18[idx_pct2_18], row_18[j] = val_j, val15
                    swapped = True
                    break
            # Jika tidak ada nilai persen yang ditemukan, set ke "-" hanya jika nilai saat ini memang bukan persen
            if not swapped and not _looks_like_percentage_value(val15):
                row_18[idx_pct2_18] = "-"
            data_18[idx_row] = row_18  # Simpan perubahan

        # Koreksi Perubahan - index 16
        val16 = get_18(idx_perubahan_18, row_18)
        is_text_16 = (_looks_like_text_not_number(val16) or _looks_like_person_name(val16) or 
                     _looks_like_securities_name(val16) or
                     (val16 != "-" and not _looks_like_percentage_value(val16) and 
                      not _looks_like_change_value(val16) and not _looks_like_large_number(val16) and len(val16) > 3))
        if is_text_16:
            swapped = False
            # Cari di blok numerik (11-17)
            for j in range(11, TARGET_COLS_18):
                if j == idx_perubahan_18:
                    continue
                v = get_18(j, row_18)
                if _looks_like_change_value(v) and not _looks_like_large_number(v) and not _looks_like_percentage_value(v):
                    row_18[idx_perubahan_18], row_18[j] = v, val16
                    swapped = True
                    break
            if not swapped or _looks_like_person_name(val16) or _looks_like_securities_name(val16):
                row_18[idx_perubahan_18] = "-"
            data_18[idx_row] = row_18  # Simpan perubahan

        # Perubahan tidak boleh berisi nilai yang sama dengan Persentase (2) (nilai persen).
        # Contoh: no 318 bawah — Perubahan salah berisi 11.76 yang seharusnya hanya di kolom (2).
        val16_after = get_18(idx_perubahan_18, row_18)
        val_pct2_18 = get_18(idx_pct2_18, row_18)
        if (val16_after != "-" and val_pct2_18 != "-" and
            _looks_like_percentage_value(val16_after) and
            str(val16_after).strip() == str(val_pct2_18).strip()):
            row_18[idx_perubahan_18] = "-"
            data_18[idx_row] = row_18

        # Perubahan tidak boleh berisi satu huruf (mis. "L" dari kolom Status yang salah tempat)
        v16 = get_18(idx_perubahan_18, row_18)
        if v16 and len(v16.strip()) == 1 and v16.strip().isalpha():
            row_18[idx_perubahan_18] = "-"
            data_18[idx_row] = row_18

    # KOREKSI BARIS KEMBAR NO: Hanya ketika nilai di (1) baris atas sama dengan (2) baris bawah (salah kolom),
    # baru pindahkan: atas (1)=bawah(1), atas (2)=atas(1). Jangan overwrite 318 atas (11.74, 11.76) dengan bawah (11.14, 11.78).
    idx_no_18 = 0
    for i in range(len(data_18) - 1):
        row_upper = data_18[i]
        row_lower = data_18[i + 1]
        no_upper = get_18(idx_no_18, row_upper)
        no_lower = get_18(idx_no_18, row_lower)
        if no_upper != no_lower or not no_upper or no_upper == "-":
            continue
        pct1_upper = get_18(idx_pct1_18, row_upper)
        pct2_upper = get_18(idx_pct2_18, row_upper)
        pct1_lower = get_18(idx_pct1_18, row_lower)
        pct2_lower = get_18(idx_pct2_18, row_lower)
        if (_looks_like_percentage_value(pct1_upper) and (pct2_upper == "-" or not pct2_upper) and
            _looks_like_percentage_value(pct1_lower) and _looks_like_percentage_value(pct2_lower) and
            str(pct1_upper).strip() == str(pct2_lower).strip()):
            row_upper[idx_pct1_18] = pct1_lower
            row_upper[idx_pct2_18] = pct1_upper
            data_18[i] = row_upper

    # Kolom Persentase (1) dan (2) merge cell: jika baris atas kosong ("-"), isi dari baris bawah (nilai sama untuk kedua baris, mis. 318)
    for i in range(len(data_18) - 1):
        row_upper = data_18[i]
        row_lower = data_18[i + 1]
        no_upper = get_18(idx_no_18, row_upper)
        no_lower = get_18(idx_no_18, row_lower)
        if no_upper != no_lower or not no_upper or no_upper == "-":
            continue
        pct1_u = get_18(idx_pct1_18, row_upper)
        pct2_u = get_18(idx_pct2_18, row_upper)
        pct1_l = get_18(idx_pct1_18, row_lower)
        pct2_l = get_18(idx_pct2_18, row_lower)
        if (pct1_u == "-" or not pct1_u) and _looks_like_percentage_value(pct1_l):
            row_upper[idx_pct1_18] = pct1_l
            data_18[i] = row_upper
        if (pct2_u == "-" or not pct2_u) and _looks_like_percentage_value(pct2_l):
            row_upper[idx_pct2_18] = pct2_l
            data_18[i] = row_upper

    # Baris bawah (No sama): isi Jumlah Saham (1)/(2) dan Perubahan dari nilai yang ada di baris itu atau baris atas (salah kolom)
    for i in range(len(data_18) - 1):
        row_upper = data_18[i]
        row_lower = data_18[i + 1]
        no_upper = get_18(idx_no_18, row_upper)
        no_lower = get_18(idx_no_18, row_lower)
        if no_upper != no_lower or not no_upper or no_upper == "-":
            continue
        j1_l = get_18(idx_jumlah1_18, row_lower)
        j2_l = get_18(idx_jumlah2_18, row_lower)
        pl = get_18(idx_perubahan_18, row_lower)
        j1_u = get_18(idx_jumlah1_18, row_upper)
        j2_u = get_18(idx_jumlah2_18, row_upper)
        sg1_l = get_18(idx_saham_gab1_18, row_lower)
        sg2_l = get_18(idx_saham_gab2_18, row_lower)
        sg1_u = get_18(idx_saham_gab1_18, row_upper)
        sg2_u = get_18(idx_saham_gab2_18, row_upper)
        # Pass awal: isi Jumlah (1)/(2) baris bawah dari Saham Gabungan baris bawah.
        # Jangan isi jika SG baris bawah sama dengan Jumlah baris atas (No 318: SG=217,622,500 = Jumlah(2) atas → biarkan untuk large_vals 2,000,000).
        both_upper_empty = (not j1_u or j1_u == "-") and (not j2_u or j2_u == "-")
        if (j1_l == "-" or not j1_l) and sg1_l and sg1_l != "-" and _looks_like_large_number(sg1_l):
            if (sg1_l == j2_u or sg1_l == j1_u):
                pass
            elif (j1_u and j1_u != "-" and sg1_l != sg1_u) or both_upper_empty:
                row_lower[idx_jumlah1_18] = sg1_l
                j1_l = sg1_l
                data_18[i + 1] = row_lower
        if (j2_l == "-" or not j2_l) and sg2_l and sg2_l != "-" and _looks_like_large_number(sg2_l):
            if (sg2_l == j1_u or sg2_l == j2_u):
                pass
            elif (j2_u and j2_u != "-" and sg2_l != sg2_u) or both_upper_empty:
                row_lower[idx_jumlah2_18] = sg2_l
                j2_l = sg2_l
                data_18[i + 1] = row_lower
        if (j1_l == "-" or not j1_l) and sg2_l and sg2_l != "-" and _looks_like_large_number(sg2_l):
            if (sg2_l == j2_u or sg2_l == j1_u):
                pass
            elif (j1_u and j1_u != "-" and sg2_l != sg2_u) or both_upper_empty:
                row_lower[idx_jumlah1_18] = sg2_l
                j1_l = sg2_l
                data_18[i + 1] = row_lower
        if (j2_l == "-" or not j2_l) and sg1_l and sg1_l != "-" and _looks_like_large_number(sg1_l):
            if (sg1_l == j1_u or sg1_l == j2_u):
                pass
            elif (j2_u and j2_u != "-" and sg1_l != sg1_u) or both_upper_empty:
                row_lower[idx_jumlah2_18] = sg1_l
                j2_l = sg1_l
                data_18[i + 1] = row_lower
        need_j1 = j1_l == "-" or not j1_l
        need_j2 = j2_l == "-" or not j2_l
        need_p = pl == "-" or not pl
        if not (need_j1 or need_j2 or need_p):
            continue
        # Kumpulkan nilai angka besar dan perubahan: dari baris bawah, lalu baris berikutnya (sama No), lalu baris atas.
        # Penting: untuk No 318, nilai 2,000,000 bisa ada di baris ke-3 (identity + 2,000,000 2,000,000 0); baris ke-2 punya SG/Persen tapi Jumlah "-".
        # Jadi kita ambil juga dari row i+2, i+3, ... selama No sama atau kosong (lanjutan).
        large_vals_by_col = []
        change_vals = []
        rows_to_scan = [(row_lower, True)]
        k = i + 2
        while k < len(data_18):
            next_row = data_18[k]
            next_no = get_18(idx_no_18, next_row)
            if next_no and next_no != "-" and next_no != no_upper:
                break
            rows_to_scan.append((next_row, True))
            k += 1
        rows_to_scan.append((row_upper, False))
        for row_src, is_lower in rows_to_scan:
            for c in range(TARGET_COLS_18):
                if not is_lower and (c == idx_jumlah1_18 or c == idx_jumlah2_18):
                    continue
                v = get_18(c, row_src)
                if v == "-" or not v:
                    continue
                if _looks_like_large_number(v) and not _looks_like_percentage_value(v):
                    if not is_lower and (j1_u and j1_u != "-") and v in (j1_u, j2_u):
                        continue
                    large_vals_by_col.append((c, v, is_lower))
                if c != idx_perubahan_18 and _looks_like_change_value(v) and not _looks_like_large_number(v):
                    change_vals.append(v)
        # Urutkan: baris bawah dulu (termasuk baris lanjutan sama No), kolom 0-10 dulu, lalu 11-17
        large_vals_by_col.sort(key=lambda x: (0 if x[2] else 1, 0 if x[0] <= 10 else 1, x[0]))
        large_vals = [v for _, v, _ in large_vals_by_col]
        sg1_u = get_18(idx_saham_gab1_18, row_upper)
        sg2_u = get_18(idx_saham_gab2_18, row_upper)
        # Kandidat Jumlah Saham: beda dari baris atas; untuk (2) saja boleh sama dengan j2_u (247 bawah isi dari nilai yang ada)
        def _ok_candidate_j1(v):
            return v not in (j1_u, j2_u, sg1_u, sg2_u)
        def _ok_candidate_j2(v):
            if v in (j1_u, sg1_u, sg2_u):
                return False
            if v == j2_u:
                return True
            return True
        # 247 bawah: jika baris atas Jumlah Saham (1) = "-", baris bawah (1) harus "-"
        if j1_u == "-":
            row_lower[idx_jumlah1_18] = "-"
            data_18[i + 1] = row_lower
            need_j1 = False
        if need_j1 or need_j2:
            candidate = None
            # Jika kita masih butuh J1, pilih kandidat J1 dulu
            if need_j1:
                for v in large_vals:
                    if _ok_candidate_j1(v):
                        candidate = v
                        break
            # Untuk J2 (termasuk kasus hanya butuh J2, seperti 247 bawah), pilih kandidat J2 langsung jika belum ada
            if need_j2 and candidate is None:
                for v in large_vals:
                    if _ok_candidate_j2(v):
                        candidate = v
                        break
            if candidate is not None:
                if need_j1:
                    row_lower[idx_jumlah1_18] = candidate
                if need_j2:
                    # Jika J1 juga diisi dengan candidate, dan tersedia nilai besar lain yang cocok untuk J2, pakai nilai lain itu.
                    # Jika hanya J2 yang dibutuhkan (J1 tidak diisi, mis. 247), langsung pakai candidate untuk J2.
                    if need_j1 and len(large_vals) >= 2:
                        other = next((x for x in large_vals if x != candidate and _ok_candidate_j2(x)), candidate)
                        row_lower[idx_jumlah2_18] = other
                    else:
                        row_lower[idx_jumlah2_18] = candidate
                data_18[i + 1] = row_lower
        # Fallback Jumlah Saham (2) dari baris atas: hanya jika baris atas J1 terisi (bukan kasus 247).
        # Untuk 247 (baris atas J1="-") J2 bawah jangan diisi dari j2_u/sg2_u agar tidak duplikat Saham Gabungan (2).
        if need_j2 and (get_18(idx_jumlah2_18, row_lower) == "-" or not get_18(idx_jumlah2_18, row_lower)):
            if (j1_u and j1_u != "-") and j2_u and j2_u != "-" and _looks_like_large_number(j2_u):
                row_lower[idx_jumlah2_18] = j2_u
                data_18[i + 1] = row_lower
            elif (j1_u and j1_u != "-") and sg2_u and sg2_u != "-" and _looks_like_large_number(sg2_u):
                row_lower[idx_jumlah2_18] = sg2_u
                data_18[i + 1] = row_lower
        # Perubahan: isi dari angka yang bukan No baris; jika tidak ada kandidat valid, isi "0" (mis. 318 bawah)
        # Jika baris atas Perubahan "-", jangan isi baris bawah dengan "0" — biarkan "-" (mis. 247)
        if need_p:
            filled = False
            no_lower_str = str(no_lower).strip() if no_lower not in (None, "") else ""
            for v in change_vals:
                vstrip = v.strip()
                if no_lower_str and vstrip == no_lower_str:
                    continue
                if len(vstrip) <= 4 and vstrip.isdigit():
                    row_lower[idx_perubahan_18] = vstrip
                    data_18[i + 1] = row_lower
                    filled = True
                    break
            pl_upper = get_18(idx_perubahan_18, row_upper)
            upper_perubahan_empty = not pl_upper or pl_upper.strip() == "-"
            if not filled and (get_18(idx_perubahan_18, row_lower) == "-" or not get_18(idx_perubahan_18, row_lower)):
                if not upper_perubahan_empty:
                    row_lower[idx_perubahan_18] = "0"
                    data_18[i + 1] = row_lower
        # Koreksi: baris bawah (No sama) yang Perubahan-nya satu huruf (mis. "D" dari kolom Domestik) → "0"
        pl_now = get_18(idx_perubahan_18, row_lower)
        if pl_now and len(pl_now.strip()) == 1 and pl_now.strip().isalpha():
            row_lower[idx_perubahan_18] = "0"
            data_18[i + 1] = row_lower

    # Pass akhir: isi Jumlah Saham (1)/(2) baris bawah dari baris mana pun dalam grup No yang sama.
    # Menangani kasus 318 ketika nilai 2,000,000 ada di baris ke-3 atau di kolom yang salah di baris ke-2.
    idx_no_18 = 0
    no_to_indices = {}
    for idx, row in enumerate(data_18):
        no = get_18(idx_no_18, row)
        if no and no != "-":
            no_to_indices.setdefault(no, []).append(idx)
        elif idx > 0:
            prev_no = get_18(idx_no_18, data_18[idx - 1])
            if prev_no and prev_no != "-":
                no_to_indices.setdefault(prev_no, []).append(idx)
    for no, indices in no_to_indices.items():
        if len(indices) < 2:
            continue
        first_idx = indices[0]
        first_row = data_18[first_idx]
        j1_first = get_18(idx_jumlah1_18, first_row)
        j2_first = get_18(idx_jumlah2_18, first_row)
        sg1_first = get_18(idx_saham_gab1_18, first_row)
        sg2_first = get_18(idx_saham_gab2_18, first_row)
        for idx in indices[1:]:
            row = data_18[idx]
            j1 = get_18(idx_jumlah1_18, row)
            j2 = get_18(idx_jumlah2_18, row)
            # Baris bawah yang duplikat (nilai sama dengan baris atas) juga harus dikoreksi ke nilai sendiri (318: 2,000,000)
            is_duplicate_of_first = (j1 == j1_first and j2 == j2_first) and (j1_first != "-" or j2_first != "-")
            need_fill_j1 = (j1 == "-" or not j1) or is_duplicate_of_first
            need_fill_j2 = (j2 == "-" or not j2) or is_duplicate_of_first
            if not need_fill_j1 and not need_fill_j2:
                continue
            large_in_group = []
            for i in indices:
                r = data_18[i]
                for c in range(TARGET_COLS_18):
                    v = get_18(c, r)
                    if v == "-" or not v:
                        continue
                    if _looks_like_large_number(v) and not _looks_like_percentage_value(v):
                        large_in_group.append(v)
            large_in_group = list(dict.fromkeys(large_in_group))
            candidate_j1 = next((v for v in large_in_group if v not in (j1_first, j2_first)), None)
            if candidate_j1 and need_fill_j1:
                row[idx_jumlah1_18] = candidate_j1
                data_18[idx] = row
                j1 = candidate_j1
            if need_fill_j2:
                if j1_first == "-" and j2_first and j2_first in large_in_group and not is_duplicate_of_first:
                    candidate_j2 = j2_first
                else:
                    candidate_j2 = next((v for v in large_in_group if v not in (j1_first, j2_first, sg1_first, sg2_first)), None)
                if candidate_j2 is None and candidate_j1:
                    candidate_j2 = candidate_j1
                if candidate_j2:
                    row[idx_jumlah2_18] = candidate_j2
                    data_18[idx] = row

    return {
        "header_top": header_top,
        "header_row": header_18,
        "data": data_18,
        "raw_data": raw_data_for_debug,
    }


def build_table_from_spans(span_items: list[dict]) -> list[list[str]]:
    """Dari daftar span dengan bbox, bangun tabel: list of rows, tiap row = list of cell text."""
    if not span_items:
        return []
    # Urutkan: halaman -> y -> x
    def key(s):
        p = s.get("page", 1)
        b = s.get("bbox") or (0, 0, 0, 0)
        y = (b[1] + b[3]) / 2
        x = b[0]
        return (p, y, x)

    sorted_spans = sorted(span_items, key=key)
    rows = []
    current_row_y = None
    current_row_cells = []
    current_cell_texts = []
    last_x1 = None

    for s in sorted_spans:
        text = (s.get("text") or "").strip()
        if not text:
            continue
        bbox = s.get("bbox") or (0, 0, 0, 0)
        x0, y0, x1, y1 = bbox
        mid_y = (y0 + y1) / 2

        # Baris baru jika y beda cukup jauh
        if current_row_y is not None and abs(mid_y - current_row_y) > ROW_Y_TOLERANCE:
            if current_cell_texts:
                current_row_cells.append(" ".join(current_cell_texts))
            if current_row_cells:
                rows.append(current_row_cells)
            current_row_cells = []
            current_cell_texts = []
            last_x1 = None

        current_row_y = mid_y

        # Kolom baru jika jarak x cukup besar
        if last_x1 is not None and (x0 - last_x1) > COLUMN_X_GAP:
            if current_cell_texts:
                current_row_cells.append(" ".join(current_cell_texts))
            current_cell_texts = []
        current_cell_texts.append(text)
        last_x1 = x1

    if current_cell_texts:
        current_row_cells.append(" ".join(current_cell_texts))
    if current_row_cells:
        rows.append(current_row_cells)
    return rows


# Opsi format output PDF: spasi antar baris & antar paragraf
OUTPUT_STYLES = {
    "paragraph": {"line_height": 12, "line_gap": 1, "para_gap": 4},   # standar, enak dibaca
    "compact": {"line_height": 10, "line_gap": 0, "para_gap": 1},      # dempet, cocok list/tabel
    "lines": {"line_height": 12, "line_gap": 1, "para_gap": 2},        # baris per baris, cocok tabel
}


# Spasi seperti Shift+Enter (baris menempel): tinggi baris = size * multiplier
TIGHT_LINE_MULT = 1.05  # sangat ketat, baris nyaris rapat


def create_pdf_raw_blue_one_per_line(lines: list[str], output_path: str) -> None:
    """Buat PDF berisi teks biru mentah: satu kata/baris, untuk debugging."""
    if not lines:
        return
    doc = fitz.open()
    blue_pdf = (0, 0, 1)
    margin = 50
    y = margin
    page_width = 595
    page_height = 842
    fontsize = 10
    line_step = fontsize * 1.4
    page = doc.new_page(width=page_width, height=page_height)
    for line in lines:
        line_safe = (line or "").strip()
        if not line_safe:
            continue
        line_safe = "".join(c if ord(c) < 256 else "?" for c in line_safe)
        pt = fitz.Point(margin, y + fontsize * 0.9)
        try:
            page.insert_text(pt, line_safe, fontsize=fontsize, color=blue_pdf, fontname="helv")
        except Exception:
            page.insert_text(pt, line_safe, fontsize=fontsize, color=blue_pdf)
        y += line_step
        if y > page_height - margin - line_step:
            page = doc.new_page(width=page_width, height=page_height)
            y = margin
    doc.save(output_path, garbage=1, deflate=False)
    doc.close()


def create_pdf_with_blue_text(
    blue_spans: list[dict], output_path: str, output_style: str = "paragraph"
) -> None:
    """Buat PDF baru yang hanya berisi teks biru (tetap warna biru).
    output_style: 'paragraph' (standar), 'compact' (rapat), 'lines' (baris per baris untuk tabel).
    Jika semua teks dari satu halaman, spasi otomatis sangat ketat (single spacing).
    """
    pages_used = {item.get("page", 1) for item in blue_spans}
    single_page = len(pages_used) <= 1
    # Compact atau satu halaman → spasi ketat seperti Shift+Enter (baris menempel)
    use_tight_spacing = single_page or (output_style == "compact")

    style = OUTPUT_STYLES.get(output_style, OUTPUT_STYLES["paragraph"])
    line_height = style["line_height"]
    line_gap = style["line_gap"]
    para_gap = style["para_gap"]
    if use_tight_spacing:
        line_gap = 0
        para_gap = 0

    doc = fitz.open()
    blue_pdf = (0, 0, 1)
    margin = 50
    y = margin
    page_width = 595
    page_height = 842
    page = doc.new_page(width=page_width, height=page_height)
    prev_source_page = None  # halaman sumber item sebelumnya
    for item in blue_spans:
        text = item.get("text") or ""
        if not text.strip():
            continue
        try:
            size = min(float(item.get("size", 12)), 12)
        except (TypeError, ValueError):
            size = 12
        # Ketat (Shift+Enter): tinggi baris = size * 1.05 agar baris menempel
        if use_tight_spacing:
            line_step = size * TIGHT_LINE_MULT
            empty_line_step = size * TIGHT_LINE_MULT * 0.4
        else:
            line_step = line_height + line_gap
            empty_line_step = line_height * 0.5

        item_page = item.get("page", 1)
        # Satu spasi antar halaman sumber: [hal 4] ... [hal 5] diberi jarak
        if prev_source_page is not None and item_page != prev_source_page:
            y += line_step
            min_line = size * (TIGHT_LINE_MULT + 0.3) if use_tight_spacing else line_height
            if y > page_height - margin - min_line:
                page = doc.new_page(width=page_width, height=page_height)
                y = margin
        prev_source_page = item_page

        label = f"[hal {item_page}] "
        full = label + text
        for line in full.split("\n"):
            line = line.strip()
            if not line:
                y += empty_line_step
                continue
            # Pastikan hanya karakter yang aman untuk font helv (Latin)
            line_safe = "".join(c if ord(c) < 256 else "?" for c in line)
            pt = fitz.Point(margin, y + size * 0.9)
            try:
                page.insert_text(pt, line_safe, fontsize=size, color=blue_pdf, fontname="helv")
            except Exception:
                page.insert_text(pt, line_safe, fontsize=size, color=blue_pdf)
            y += line_step
            min_line = size * (TIGHT_LINE_MULT + 0.3) if use_tight_spacing else line_height
            if y > page_height - margin - min_line:
                page = doc.new_page(width=page_width, height=page_height)
                y = margin
        y += para_gap
        min_line = size * (TIGHT_LINE_MULT + 0.3) if use_tight_spacing else line_height
        if y > page_height - margin - min_line:
            page = doc.new_page(width=page_width, height=page_height)
            y = margin
    doc.save(output_path, garbage=1, deflate=False)
    doc.close()


def create_pdf_from_table(table: list[list[str]], output_path: str) -> None:
    """Buat PDF dengan tabel: grid garis + teks biru di tiap sel."""
    if not table:
        return
    blue_pdf = (0, 0, 1)
    margin = 40
    page_width = 595
    page_height = 842
    fontsize = 9
    cell_pad = 4
    row_height = fontsize * 1.4 + cell_pad * 2

    num_cols = max(len(row) for row in table) if table else 0
    if num_cols == 0:
        return
    # Normalisasi: setiap baris punya num_cols sel
    rows = [list(row) + [""] * (num_cols - len(row)) for row in table]
    # Perkiraan lebar kolom: bagi rata area konten
    content_width = page_width - 2 * margin
    col_width = content_width / num_cols

    doc = fitz.open()
    page = doc.new_page(width=page_width, height=page_height)
    y = margin

    for r_idx, row in enumerate(rows):
        if y + row_height > page_height - margin:
            page = doc.new_page(width=page_width, height=page_height)
            y = margin
        x = margin
        for c_idx, cell_text in enumerate(row):
            text_safe = "".join(c if ord(c) < 256 else "?" for c in (cell_text or ""))
            rect = fitz.Rect(x, y, x + col_width, y + row_height)
            # Garis batas sel
            page.draw_rect(rect, color=(0.7, 0.7, 0.7), width=0.5)
            # Teks di dalam sel (clip agar tidak keluar)
            try:
                page.insert_textbox(
                    fitz.Rect(x + cell_pad, y + cell_pad, x + col_width - cell_pad, y + row_height - cell_pad),
                    text_safe,
                    fontsize=fontsize,
                    color=blue_pdf,
                    fontname="helv",
                    align=fitz.TEXT_ALIGN_LEFT,
                )
            except Exception:
                page.insert_text(
                    fitz.Point(x + cell_pad, y + cell_pad + fontsize * 0.9),
                    text_safe[:100],
                    fontsize=fontsize,
                    color=blue_pdf,
                )
            x += col_width
        y += row_height
    doc.save(output_path, garbage=1, deflate=False)
    doc.close()


def _column_index_by_header(header_row: list, name_keywords: tuple) -> int:
    """Cari indeks kolom yang judulnya mengandung salah satu keyword (case-insensitive). Return -1 jika tidak ketemu."""
    for idx, h in enumerate(header_row):
        if not h:
            continue
        h_lower = str(h).strip().lower()
        for kw in name_keywords:
            if kw.lower() in h_lower:
                return idx
    return -1


def _apply_raw_blue_fix_same_no_baris_bawah(
    data_rows: list,
    raw_blue_lines: list,
    header_row: list | None = None,
    debug_ref: dict | None = None,
) -> None:
    """
    Isi Jumlah Saham (1)/(2) dan Perubahan baris bawah HANYA dari urutan raw teks biru (bukan dari Saham Gabungan).
    Berlaku untuk SEMUA No yang punya lebih dari satu baris (tidak hardcode nomor tertentu).
    Untuk setiap No demikian: cari di raw_blue_lines di segmen setelah kemunculan No tersebut
    sebuah triple (angka_besar, angka_besar, perubahan) yang berbeda dari baris pertama,
    lalu isi baris kedua dengan triple itu (prioritas triple dengan Perubahan=0). SG (1)/(2) baris bawah = merge cell (sama dengan baris atas).
    Jika debug_ref diberikan, isi debug_ref["707"] saat No 707 diproses (untuk logging frontend).
    """
    if not data_rows or not raw_blue_lines or len(data_rows) < 2:
        return
    ncols = max(len(r) for r in data_rows)
    if ncols < 12:
        return

    # Indeks kolom: pakai nama header jika ada, fallback ke posisi tetap (sesuai TEMPLATE_HEADER_18)
    if header_row and len(header_row) >= ncols:
        idx_no = _column_index_by_header(header_row, ("no", "no."))
        idx_j1 = _column_index_by_header(header_row, ("jumlah saham (1)", "jumlah saham(1)"))
        idx_j2 = _column_index_by_header(header_row, ("jumlah saham (2)", "jumlah saham(2)"))
        idx_perubahan = _column_index_by_header(header_row, ("perubahan",))
        idx_saham_gab1 = _column_index_by_header(header_row, ("saham gabungan per investor (1)", "saham gabungan (1)"))
        idx_saham_gab2 = _column_index_by_header(header_row, ("saham gabungan per investor (2)", "saham gabungan (2)"))
        if idx_no < 0:
            idx_no = 0
        if idx_j1 < 0:
            idx_j1 = 11
        if idx_j2 < 0:
            idx_j2 = 14
        if idx_perubahan < 0:
            idx_perubahan = 17
        if idx_saham_gab1 < 0:
            idx_saham_gab1 = 12
        if idx_saham_gab2 < 0:
            idx_saham_gab2 = 15
    else:
        idx_no, idx_j1, idx_j2, idx_perubahan = 0, 11, 14, 17
        idx_saham_gab1, idx_saham_gab2 = 12, 15

    if ncols <= max(idx_j1, idx_j2, idx_perubahan):
        return
    lines = [str(s).strip() for s in raw_blue_lines]

    # Normalisasi baris dan kelompokkan per No
    no_to_indices = {}
    for row_idx, row in enumerate(data_rows):
        row = list(row)
        while len(row) < ncols:
            row.append("-")
        data_rows[row_idx] = row
        no = (row[idx_no] or "").strip()
        if no and no != "-":
            no_to_indices.setdefault(no, []).append(row_idx)
        elif row_idx > 0:
            prev_no = (data_rows[row_idx - 1][idx_no] or "").strip()
            if prev_no and prev_no != "-":
                no_to_indices.setdefault(prev_no, []).append(row_idx)

    # Proses setiap No yang punya 2+ baris (generik; tidak hardcode 318 atau nomor lain)
    for no, indices in no_to_indices.items():
        if len(indices) < 2:
            continue
        first_idx, second_idx = indices[0], indices[1]
        first_row = data_rows[first_idx]
        j1_first = (first_row[idx_j1] or "").strip() or "-"
        j2_first = (first_row[idx_j2] or "").strip() or "-"

        # Posisi pertama kemunculan No ini di raw
        pos_no = None
        for i, w in enumerate(lines):
            if (w or "").strip() == no:
                pos_no = i
                break
        if pos_no is None:
            if debug_ref is not None and no == "707":
                debug_ref["707"] = {"path": "pos_no is None", "no_in_raw": no in lines, "sample": lines[:20]}
            continue

        # Batas: sampai kemunculan No lain (nomor urut berikutnya) atau akhir list
        segment_end = len(lines)
        next_no_str = None
        for j in range(pos_no + 1, len(lines)):
            w = (lines[j] or "").strip()
            if _looks_like_no(w) and w != no:
                segment_end = j
                next_no_str = w
                break

        # Cari triple (angka_besar, angka_besar, perubahan) di segmen yang berbeda dari baris pertama.
        # Untuk baris bawah (mis. 318 bawah): pilih triple dengan Perubahan=0 jika ada, else triple terakhir.
        # Rentang: i sampai i+2 harus dalam segment; sertakan sampai ujung segmen.
        seg_end = min(segment_end, len(lines))
        i_end = min(seg_end - 1, len(lines) - 2) if len(lines) >= 2 else pos_no + 1
        candidates = []
        for i in range(pos_no + 1, i_end):
            v1, v2, v3 = (lines[i] or "").strip(), (lines[i + 1] or "").strip(), (lines[i + 2] or "").strip()
            ok1 = _looks_like_large_number(v1) and not _looks_like_percentage_value(v1)
            ok2 = _looks_like_large_number(v2) and not _looks_like_percentage_value(v2)
            ok3 = _looks_like_change_value(v3)
            if not ok1 or not ok2 or not ok3:
                continue
            if (v1, v2) != (j1_first, j2_first):
                candidates.append((i, v3))
        triple_start = None
        for i, pv in candidates:
            if (pv or "").strip() == "0":
                triple_start = i
                break
        if triple_start is None and candidates:
            triple_start = candidates[-1][0]
        if triple_start is None:
            # Fallback: baris bawah hanya punya satu nilai (mis. 487 di Jumlah Saham (2)), baris atas Perubahan = -nilai
            # Hanya nilai 3+ digit (100-9999) agar tidak ambil "2", "19" dari alamat (Lantai 2, Ruang 210, dll.)
            def _collect_change_like(start: int, end: int):
                out_list = []
                for i in range(start, end):
                    w = (lines[i] or "").strip()
                    if not w or w == "-":
                        continue
                    w_normalized = w.replace(",", "").replace(" ", "").replace(".", "", 1)
                    if not w_normalized.isdigit():
                        continue
                    try:
                        v = int(w_normalized)
                    except ValueError:
                        continue
                    if _looks_like_large_number(w) or _looks_like_percentage_value(w):
                        continue
                    # Minimal 3 digit (100-9999) untuk Perubahan/J2 yang wajar; hindari 0, 2, 19 dari teks
                    if 100 <= v <= 9999:
                        out_list.append((i, str(v)))
                return out_list

            change_like_in_segment = _collect_change_like(pos_no + 1, seg_end)
            if not change_like_in_segment:
                seg_end_ext = min(seg_end + 80, len(lines))
                change_like_in_segment = _collect_change_like(pos_no + 1, seg_end_ext)
            # Filter generik: buang nilai yang sama dengan No baris berikutnya (dari raw)
            def _norm_num(s):
                try:
                    return str(int((s or "").replace(",", "").replace(" ", "")))
                except (ValueError, TypeError):
                    return (s or "").strip()
            if next_no_str is not None:
                next_norm = _norm_num(next_no_str)
                change_like_in_segment = [(i, v) for i, v in change_like_in_segment if _norm_num(v) != next_norm]
            # Jika 3+ nilai: ujung sering noise (alamat di awal, No berikutnya di akhir); ambil yang di tengah
            if len(change_like_in_segment) >= 3:
                change_like_in_segment = change_like_in_segment[1:-1]
            # Jika 2 nilai: biasanya [noise, nilai asli]; pakai nilai terakhir untuk keduanya
            if len(change_like_in_segment) == 2:
                change_like_in_segment = [change_like_in_segment[-1]]
            apply_fallback = False
            if change_like_in_segment:
                row2_check = data_rows[second_idx]
                j2_current = (row2_check[idx_j2] or "").strip() or "-"
                # Jangan pakai fallback jika baris bawah J2 sudah angka besar (mis. 247: 8,417,464,300); biarkan Perubahan atas "-" dan J2 bawah tetap
                if not _looks_like_large_number(j2_current):
                    apply_fallback = True
                    _, val_first = change_like_in_segment[0]
                    val_last = change_like_in_segment[-1][1] if change_like_in_segment else val_first
                    if len(change_like_in_segment) == 1:
                        val_last = val_first
                    cur_perubahan = (first_row[idx_perubahan] or "").strip() or "-"
                    if cur_perubahan == "-" or not cur_perubahan:
                        first_row[idx_perubahan] = "-" + val_first if val_first.lstrip("-") == val_first else val_first
                    row2 = list(data_rows[second_idx])
                    while len(row2) < ncols:
                        row2.append("-")
                    row2[idx_j1] = "-"
                    row2[idx_j2] = val_last
                    row2[idx_perubahan] = (row2[idx_perubahan] or "").strip() or "-"
                    if idx_saham_gab1 < ncols:
                        row2[idx_saham_gab1] = (first_row[idx_saham_gab1] or "").strip() or "-"
                    if idx_saham_gab2 < ncols:
                        row2[idx_saham_gab2] = (first_row[idx_saham_gab2] or "").strip() or "-"
                    data_rows[second_idx] = row2
                else:
                    # Baris bawah J2 sudah angka besar: kosongkan J1 hanya jika J1 = J2 (duplikat salah tempat, mis. 247)
                    # Jika J2 baris bawah = J2/SG(2) baris atas (duplikat), cari di raw angka besar lain untuk J2 bawah (mis. 247: 8,417,464,300)
                    row2 = list(data_rows[second_idx])
                    while len(row2) < ncols:
                        row2.append("-")
                    j1_current = (row2[idx_j1] or "").strip() or "-"
                    j2_val = (row2[idx_j2] or "").strip() or "-"
                    sg2_first = (first_row[idx_saham_gab2] or "").strip() or "-" if idx_saham_gab2 < len(first_row) else "-"

                    def _norm(s):
                        return (s or "").replace(",", "").replace(" ", "")

                    j2_val_norm = _norm(j2_val)
                    j2_first_norm = _norm(j2_first)
                    sg2_first_norm = _norm(sg2_first)
                    # Apakah J2 bawah saat ini sama dengan nilai baris atas (salah tempat)?
                    j2_equals_upper = j2_val_norm and (j2_val_norm == j2_first_norm or j2_val_norm == sg2_first_norm)

                    if j2_equals_upper:
                        # Cari angka besar lain untuk J2 bawah: dalam segmen, atau sedikit setelah segmen hanya jika orde sama dengan SG(2)
                        # (e.g. 8,417,464,300) agar tidak ambil 2,000,000 dari No berikutnya
                        large_in_segment = []
                        seg_end_ext = min(seg_end + 40, len(lines))
                        for ii in range(pos_no + 1, seg_end_ext):
                            w = (lines[ii] or "").strip()
                            if not w or not _looks_like_large_number(w) or _looks_like_percentage_value(w):
                                continue
                            w_norm = _norm(w)
                            if w_norm in (j2_first_norm, sg2_first_norm, _norm(j1_first)):
                                continue
                            # Di luar segmen (ii >= seg_end): hanya terima nilai yang orde sama dengan SG(2) (e.g. 8,417,xxx)
                            if ii >= seg_end and sg2_first_norm and len(sg2_first_norm) >= 4 and len(w_norm) >= 4:
                                if w_norm[:4] != sg2_first_norm[:4]:
                                    continue
                            large_in_segment.append((ii, w))
                        # Ambil nilai besar terakhir yang lolos filter (nilai baris bawah di PDF, e.g. 8,417,464,300)
                        if large_in_segment:
                            _, replace_j2 = large_in_segment[-1]
                            row2[idx_j2] = replace_j2
                            data_rows[second_idx] = row2

                    j1_eq_j2 = (j1_current == j2_val) or (
                        j1_current and j2_val
                        and _looks_like_large_number(j1_current)
                        and _looks_like_large_number(j2_val)
                        and (j1_current.replace(",", "").replace(" ", "") == j2_val.replace(",", "").replace(" ", ""))
                    )
                    if j1_current != "-" and j1_eq_j2:
                        row2[idx_j1] = "-"
                        data_rows[second_idx] = row2
            if debug_ref is not None and no == "707":
                segment_preview = lines[pos_no : seg_end]
                j2_after = (data_rows[second_idx][idx_j2] or "").strip() or "-" if second_idx < len(data_rows) else "-"
                debug_ref["707"] = {
                    "path": "fallback (no triple)",
                    "pos_no": pos_no,
                    "segment_end": segment_end,
                    "seg_end": seg_end,
                    "j1_first": j1_first,
                    "j2_first": j2_first,
                    "segment_preview": segment_preview[:80],
                    "change_like_in_segment": change_like_in_segment,
                    "fallback_applied": apply_fallback,
                    "first_row_perubahan_after": (first_row[idx_perubahan] or "").strip() or "-",
                    "second_row_j2_after": j2_after,
                }
            continue

        j1_val = lines[triple_start]
        j2_val = lines[triple_start + 1]
        p_val = lines[triple_start + 2]
        row = list(data_rows[second_idx])
        while len(row) < ncols:
            row.append("-")
        second_row_before = list(row)
        # Kasus 707: triple (large, large, small) tapi baris bawah harus J1="-", J2=nilai kecil (487), Perubahan=0
        p_val_small = (
            _looks_like_change_value(p_val)
            and not _looks_like_large_number(p_val)
            and (p_val or "").strip() not in ("", "-", "0")
        )
        j1_empty = ((row[idx_j1] or "").strip() or "-") == "-"
        j2_empty = ((row[idx_j2] or "").strip() or "-") == "-"
        if p_val_small and j1_empty and j2_empty:
            cur_perubahan = (first_row[idx_perubahan] or "").strip() or "-"
            if cur_perubahan == "-" or not cur_perubahan:
                first_row[idx_perubahan] = "-" + p_val if (p_val or "").lstrip("-") == (p_val or "") else p_val
            row[idx_j1] = "-"
            row[idx_j2] = (p_val or "").strip()
            row[idx_perubahan] = "0"
        else:
            row[idx_j1] = j1_val
            row[idx_j2] = j2_val
            row[idx_perubahan] = p_val
        # Saham Gabungan (1)/(2) baris bawah = merge cell: sama dengan baris atas
        if idx_saham_gab1 < ncols:
            row[idx_saham_gab1] = (first_row[idx_saham_gab1] or "").strip() or "-"
        if idx_saham_gab2 < ncols:
            row[idx_saham_gab2] = (first_row[idx_saham_gab2] or "").strip() or "-"
        data_rows[second_idx] = row

        if debug_ref is not None and no == "707":
            debug_ref["707"] = {
                "path": "triple",
                "triple": (j1_val, j2_val, p_val),
                "second_row_after": list(row)[: max(idx_j1, idx_j2, idx_perubahan) + 1],
            }


@app.route("/")
def index():
    return render_template("index.html")


@app.route("/api/extract-blue", methods=["POST"])
def extract_blue():
    """Ekstrak teks biru, bangun tabel dari posisi, kembalikan JSON { table } untuk ditampilkan di halaman."""
    if "file" not in request.files:
        return jsonify({"error": "Tidak ada file"}), 400
    file = request.files["file"]
    if file.filename == "":
        return jsonify({"error": "File tidak dipilih"}), 400
    if not file.filename.lower().endswith(".pdf"):
        return jsonify({"error": "Hanya file PDF yang didukung"}), 400
    tmp_in_path = None
    try:
        with tempfile.NamedTemporaryFile(delete=False, suffix=".pdf") as tmp_in:
            file.save(tmp_in.name)
            tmp_in_path = tmp_in.name
        # Bangun tabel: baris pertama = header dari PDF (teks apa saja), sisanya = teks biru per kolom
        result = build_table_with_header_from_pdf(tmp_in_path)
        if not result:
            return jsonify({
                "error": "Tidak ada teks warna biru ditemukan di PDF ini.",
                "hint": "Pastikan teks benar-benar menggunakan warna biru (bukan hitam/abu-abu)."
            }), 422
        # Raw teks biru: ambil semua span biru, satu kata satu baris (untuk debugging)
        raw_blue_lines = []
        try:
            blue_spans = extract_blue_spans_with_bbox(tmp_in_path)
            for item in blue_spans:
                text = (item.get("text") or "").strip()
                for word in text.split():
                    if word:
                        raw_blue_lines.append(word)
        except Exception:
            raw_blue_lines = []
        if isinstance(result, dict):
            data_rows = result["data"] or []
            header_row = result.get("header_row") or list(TEMPLATE_HEADER_18)
            debug_318_ref = {}
            _apply_raw_blue_fix_same_no_baris_bawah(data_rows, raw_blue_lines, header_row, debug_ref=debug_318_ref)
            target_cols = len(header_row)
            # Pastikan header dan setiap baris data punya kolom sesuai header
            header_row = (list(header_row) + [""] * target_cols)[:target_cols]
            table = [header_row]
            for row in data_rows:
                r = (list(row) + ["-"] * target_cols)[:target_cols]
                while len(r) < target_cols:
                    r.append("-")
                table.append(r[:target_cols])
            out = {
                "table": table,
                "header_top": result.get("header_top") or [],
                "raw_data": result.get("raw_data") or [],
                "raw_blue_lines": raw_blue_lines,
            }
            if debug_318_ref.get("707") is not None:
                out["debug_707"] = debug_318_ref["707"]
            return jsonify(out)
        return jsonify({"table": result})
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
    app.run(host="0.0.0.0", port=5000, debug=True)
