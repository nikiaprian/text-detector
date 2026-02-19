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
# ROW_Y_TOLERANCE kecil (2) agar baris sub-header "Kepemilikan Per 28-JAN-2026" dan
# baris header utama "No, Kode Efek, Nama Emiten..." tidak digabung jadi satu baris.
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
# Template header 18 kolom (internal); tampilan 17 kolom setelah gabung Alamat.
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


def _merge_split_kode_emiten_rows(raw_data_rows: list[tuple], num_cols: int) -> list[tuple]:
    """
    Gabungkan baris yang terpecah:
    - Pattern A: baris i kolom 1 = Nama Emiten (salah), kolom 2 kosong; baris berikut ada Kode Efek → pindah Nama Emiten ke kolom 2, isi kolom 1 dengan Kode Efek.
    - Pattern B: baris i kolom 1 = Kode Efek, kolom 2 kosong; baris berikut kolom 1 = Nama Emiten → isi kolom 2 dengan Nama Emiten dari baris berikut.
    """
    if num_cols < 3:
        return raw_data_rows
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
        # Jika tidak ada nilai persen yang ditemukan, set ke "-"
        if not swapped:
            while len(cells) <= idx_pct1:
                cells.append("-")
            cells[idx_pct1] = "-"

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
    Koreksi in-place: hanya jika nilai di Perubahan (17) berbentuk persen (mis. 5.00, 11.70)
    dan kolom Persentase (16) kosong, pindahkan ke kolom 16. Nilai bulat (343, 0) tetap di Perubahan.
    """
    if num_cols < 18:
        return
    idx_persentase = 16
    idx_perubahan = 17
    val_perubahan = (cells[idx_perubahan] if len(cells) > idx_perubahan else "").strip()
    val_persentase = (cells[idx_persentase] if len(cells) > idx_persentase else "").strip()
    if not val_perubahan or val_perubahan == "-":
        return
    if not _looks_like_percentage_value(val_perubahan):
        return
    if not val_persentase or val_persentase == "-":
        while len(cells) <= idx_persentase:
            cells.append("-")
        cells[idx_persentase] = val_perubahan
    if len(cells) > idx_perubahan:
        cells[idx_perubahan] = "-"


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
    punya Kode Efek, salin Kode Efek ke baris pertama dan buang baris kedua (rapikan duplikat).
    """
    if num_cols < 3 or not rows:
        return rows
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
    
    # Proses setiap baris: tentukan kolom dari POSISI bbox (column_index_for_span),
    # BUKAN dari urutan span. Ini mencegah salah kolom ketika ada kolom kosong di PDF.
    raw_data_rows = []
    sorted_row_keys = sorted(rows_by_cluster.keys(), key=lambda k: (k[0], k[1]))
    
    for (page, cluster_y) in sorted_row_keys:
        spans_in_row = rows_by_cluster[(page, cluster_y)]
        cells = [""] * num_cols
        for span in spans_in_row:
            text = (span.get("text") or "").strip()
            if not text:
                continue
            col_idx = column_index_for_span(span["bbox"])
            if col_idx < 0:
                col_idx = 0
            if col_idx >= num_cols:
                col_idx = num_cols - 1
            if cells[col_idx]:
                cells[col_idx] = cells[col_idx] + " " + text
            else:
                cells[col_idx] = text
        cells = [_normalize_cell(c) for c in cells]
        if any(c.strip() for c in cells):
            raw_data_rows.append((cluster_y, cells, page))
    
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

    # Gabung kolom Alamat + Alamat (Lanjutan) jadi satu kolom agar tabel tidak terlalu lebar
    header_17 = list(template_header_row[:6]) + ["Alamat"] + list(template_header_row[8:])
    data_17 = []
    for row in final_data_rows:
        alamat_merged = ((row[6] or "").strip() + " " + (row[7] or "").strip()).strip() or "-"
        row_17 = list(row[:6]) + [alamat_merged] + list(row[8:])
        data_17.append(row_17)
    
    # Pastikan semua baris punya panjang 17 kolom sebelum koreksi
    for i, row_17 in enumerate(data_17):
        while len(row_17) < 17:
            row_17.append("-")
        data_17[i] = row_17[:17]
    
    # KOREKSI SETELAH MERGE ALAMAT: Setelah merge, tabel jadi 17 kolom (bukan 18)
    # Indeks baru: Persentase(1)=12, Persentase(2)=15, Perubahan=16 (bukan 13,16,17)
    TARGET_COLS_17 = 17
    idx_pct1_17 = 12  # Sebelumnya 13, setelah merge jadi 12
    idx_pct2_17 = 15  # Sebelumnya 16, setelah merge jadi 15
    idx_perubahan_17 = 16  # Sebelumnya 17, setelah merge jadi 16
    
    def get_17(i: int, cells: list) -> str:
        return (cells[i] if i < len(cells) else "").strip() or "-"
    
    for idx_row, row_17 in enumerate(data_17):
        # Pastikan panjang 17 kolom
        while len(row_17) < TARGET_COLS_17:
            row_17.append("-")
        row_17 = row_17[:TARGET_COLS_17]
        data_17[idx_row] = row_17  # Pastikan perubahan tersimpan
        
        # Koreksi Persentase (1) - index 12
        val12 = get_17(idx_pct1_17, row_17)
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
            for j in range(TARGET_COLS_17):
                if j == idx_pct1_17 or j == idx_pct2_17:
                    continue
                val_j = get_17(j, row_17)
                if _looks_like_percentage_value(val_j):
                    row_17[idx_pct1_17], row_17[j] = val_j, val12
                    swapped = True
                    break
            # Jika tidak ada nilai persen yang ditemukan, set ke "-"
            if not swapped:
                row_17[idx_pct1_17] = "-"
            data_17[idx_row] = row_17  # Simpan perubahan
        
        # Koreksi Persentase (2) - index 15
        val15 = get_17(idx_pct2_17, row_17)
        # Deteksi lebih agresif: jika bukan persen dan bukan angka besar, anggap teks
        is_not_pct_15 = not _looks_like_percentage_value(val15)
        is_not_large_num_15 = not _looks_like_large_number(val15)
        is_text_15 = (val15 != "-" and is_not_pct_15 and is_not_large_num_15 and 
                      (_looks_like_text_not_number(val15) or _looks_like_securities_name(val15) or 
                       _looks_like_person_name(val15) or len(val15) > 3))
        if is_text_15:
            swapped = False
            # Cari nilai persen di seluruh baris (0-16)
            for j in range(TARGET_COLS_17):
                if j == idx_pct1_17 or j == idx_pct2_17:
                    continue
                val_j = get_17(j, row_17)
                if _looks_like_percentage_value(val_j):
                    row_17[idx_pct2_17], row_17[j] = val_j, val15
                    swapped = True
                    break
            # Jika tidak ada nilai persen yang ditemukan, set ke "-"
            if not swapped:
                row_17[idx_pct2_17] = "-"
            data_17[idx_row] = row_17  # Simpan perubahan
        
        # Koreksi Perubahan - index 16
        val16 = get_17(idx_perubahan_17, row_17)
        is_text_16 = (_looks_like_text_not_number(val16) or _looks_like_person_name(val16) or 
                     _looks_like_securities_name(val16) or
                     (val16 != "-" and not _looks_like_percentage_value(val16) and 
                      not _looks_like_change_value(val16) and not _looks_like_large_number(val16) and len(val16) > 3))
        if is_text_16:
            swapped = False
            # Cari di blok numerik (10-16)
            for j in range(10, TARGET_COLS_17):
                if j == idx_perubahan_17:
                    continue
                v = get_17(j, row_17)
                if _looks_like_change_value(v) and not _looks_like_large_number(v) and not _looks_like_percentage_value(v):
                    row_17[idx_perubahan_17], row_17[j] = v, val16
                    swapped = True
                    break
            if not swapped or _looks_like_person_name(val16) or _looks_like_securities_name(val16):
                row_17[idx_perubahan_17] = "-"
            data_17[idx_row] = row_17  # Simpan perubahan

    return {
        "header_top": header_top,
        "header_row": header_17,
        "data": data_17,
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
        if isinstance(result, dict):
            header_row = result["header_row"] or list(TEMPLATE_HEADER_18)
            data_rows = result["data"] or []
            target_cols = len(header_row)
            # Pastikan header dan setiap baris data punya kolom sesuai header
            header_row = (list(header_row) + [""] * target_cols)[:target_cols]
            table = [header_row]
            for row in data_rows:
                r = (list(row) + ["-"] * target_cols)[:target_cols]
                while len(r) < target_cols:
                    r.append("-")
                table.append(r[:target_cols])
            return jsonify({"table": table, "header_top": result.get("header_top") or []})
        return jsonify({"table": result})
    except Exception as e:
        return jsonify({"error": str(e)}), 500
    finally:
        if tmp_in_path and os.path.exists(tmp_in_path):
            try:
                os.unlink(tmp_in_path)
            except Exception:
                pass


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
