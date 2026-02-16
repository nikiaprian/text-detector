"""
Backend Flask: upload PDF, ekstrak teks warna biru, generate PDF output.
"""
import os
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
# Template header baris kedua (nama kolom): dipakai tetap, isi datanya dari PDF.
# Urutan: kolom identitas, Status, lalu per periode Kepemilikan 3 subkolom, terakhir Perubahan.
TEMPLATE_HEADER_FIXED = (
    "No Kode Efek",
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


def build_template_header_row(num_cols: int) -> list[str]:
    """
    Buat baris header tetap (template) untuk num_cols kolom.
    Urutan: kolom tetap (No Kode Efek ... Status), lalu blok Kepemilikan (3 subkolom per periode), lalu Perubahan.
    """
    d = max(0, (num_cols - 11) // 3)
    row = list(TEMPLATE_HEADER_FIXED) + list(KEPEMILIKAN_SUBCOLUMNS) * d + ["Perubahan"]
    if len(row) < num_cols:
        row.extend(f"Kolom {i + 1}" for i in range(len(row), num_cols))
    elif len(row) > num_cols:
        row = row[:num_cols]
    return row


def build_template_header_row_from_header_top(header_top: list[dict]) -> list[str]:
    """
    Bangun baris header (nama kolom) dari struktur header_top agar setiap blok
    'Kepemilikan Per' punya 3 subkolom: Jumlah Saham, Saham Gabungan Per Investor,
    Persentase Kepemilikan Per Investor (%).
    """
    row = []
    fixed_used = 0
    total_cols = sum(h.get("colspan", 1) for h in header_top)
    subcols = list(KEPEMILIKAN_SUBCOLUMNS)

    for i, h in enumerate(header_top):
        c = h.get("colspan", 1)
        text = (h.get("text") or "").strip().lower()

        if fixed_used < 10:
            # Kolom tetap: No Kode Efek ... Status (Lokal/Asing)
            n_fixed = min(c, 10 - fixed_used)
            row.extend(list(TEMPLATE_HEADER_FIXED)[fixed_used : fixed_used + n_fixed])
            fixed_used += n_fixed
            remaining = c - n_fixed
            if remaining <= 0:
                continue
            # Sisa span ini (setelah kolom ke-10) = blok Kepemilikan
            if "kepemilikan" in text and remaining >= 1:
                row.extend(subcols[: min(remaining, 3)])
                if remaining > 3:
                    row.extend([""] * (remaining - 3))
            else:
                row.extend([""] * remaining)
            continue

        if "kepemilikan" in text:
            # Selalu tampilkan 3 subkolom: Jumlah Saham, Saham Gabungan, Persentase
            row.extend(subcols[: min(c, 3)])
            if c > 3:
                row.extend([""] * (c - 3))
        elif "perubahan" in text or (i == len(header_top) - 1 and c == 1):
            row.append("Perubahan")
            if c > 1:
                row.extend([""] * (c - 1))
        else:
            row.extend([""] * c)

    return row[:total_cols] if len(row) >= total_cols else row + [""] * (total_cols - len(row))


def _is_merged_kepemilikan_cell(cell_text: str) -> bool:
    """
    True jika sel header berisi gabungan sub-kolom Kepemilikan (Jumlah Saham + Saham Gabungan + Persentase)
    sehingga perlu dipecah jadi 3 kolom terpisah.
    """
    if not cell_text or not isinstance(cell_text, str):
        return False
    c = cell_text.lower().strip()
    if len(c) < 15:
        return False
    has_jumlah_saham = "jumlah" in c and "saham" in c
    has_gabungan = "gabungan" in c or ("saham" in c and "investor" in c)
    has_persentase = "persentase" in c or "kepemilikan" in c
    # Minimal 2 dari 3 sub-kolom ada dalam satu sel = dianggap merged
    return sum([has_jumlah_saham, has_gabungan, has_persentase]) >= 2


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
    Baca header tabel dari PDF dengan deteksi kata kunci: baris yang berisi
    "No", "Kode Efek", "Nama Emiten", "Nama Pemegang Rekening Efek", "Nama Pemegang Saham"
    (minimal 2 cocok) dipakai sebagai header. Isi data dari teks biru ke kolom yang sesuai.
    """
    all_spans = extract_all_spans_with_bbox(input_path)
    if not all_spans:
        return []
    rows_raw = _group_spans_into_rows(all_spans)
    if not rows_raw:
        return []

    # Cari baris pertama yang teksnya mirip header tabel (trigger kata kunci)
    header_row_idx = None
    for i, (_y, _page, row_spans) in enumerate(rows_raw):
        if _row_looks_like_header(row_spans):
            header_row_idx = i
            break
    if header_row_idx is None:
        # Fallback: tidak ada baris yang cocok kata kunci, pakai logika lama (hanya biru)
        blue_only = [s for s in all_spans if s.get("is_blue")]
        return build_table_from_spans(blue_only)

    header_spans = rows_raw[header_row_idx][2]

    # Bentuk kolom: gabung span yang berdekatan (gap kecil) jadi satu sel header
    column_boundaries = []
    header_cells = []
    cell_texts = []
    cell_x0 = cell_x1 = None
    for s in header_spans:
        bbox = s.get("bbox") or (0, 0, 0, 0)
        x0, x1 = bbox[0], bbox[2]
        if cell_x1 is not None and (x0 - cell_x1) > COLUMN_X_GAP:
            if cell_texts:
                header_cells.append(" ".join(cell_texts))
                column_boundaries.append((cell_x0, cell_x1))
            cell_texts = []
            cell_x0 = cell_x1 = None
        cell_texts.append(s.get("text") or "")
        if cell_x0 is None:
            cell_x0 = x0
        cell_x1 = x1
    if cell_texts:
        header_cells.append(" ".join(cell_texts))
        column_boundaries.append((cell_x0, cell_x1))

    if not column_boundaries:
        blue_only = [s for s in all_spans if s.get("is_blue")]
        return build_table_from_spans(blue_only)

    # Pecah sel header yang menggabungkan "Jumlah Saham" + "Saham Gabungan" + "Persentase" jadi 3 kolom
    new_boundaries = []
    new_header_cells = []
    for j, (cx0, cx1) in enumerate(column_boundaries):
        cell = header_cells[j] if j < len(header_cells) else ""
        if _is_merged_kepemilikan_cell(cell):
            w = (cx1 - cx0) / 3
            new_boundaries.append((cx0, cx0 + w))
            new_boundaries.append((cx0 + w, cx0 + 2 * w))
            new_boundaries.append((cx0 + 2 * w, cx1))
            new_header_cells.extend(KEPEMILIKAN_SUBCOLUMNS)
        else:
            new_boundaries.append((cx0, cx1))
            new_header_cells.append(cell)
    column_boundaries = new_boundaries
    header_cells = new_header_cells

    num_cols = len(column_boundaries)
    # Hilangkan celah antar kolom: bagi wilayah antara batas kiri/kanan setiap dua kolom
    # agar teks yang jatuh di celah tidak salah masuk ke kolom kiri.
    x_min_all = column_boundaries[0][0]
    x_max_all = column_boundaries[-1][1]
    new_boundaries = []
    for j in range(num_cols):
        cx0, cx1 = column_boundaries[j]
        if j == 0:
            start = x_min_all
        else:
            prev_end = column_boundaries[j - 1][1]
            start = (prev_end + cx0) / 2
        if j == num_cols - 1:
            end = x_max_all
        else:
            next_start = column_boundaries[j + 1][0]
            end = (cx1 + next_start) / 2
        new_boundaries.append((start, end))
    column_boundaries = new_boundaries

    def column_index_for_span(bbox) -> int:
        x0, _, x1, _ = bbox
        mid_x = (x0 + x1) / 2
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

    # Baris header atas (Kepemilikan Per 28-JAN-2026, Kepemilikan Per 29-JAN-2026, dll) untuk 2 baris header
    header_top = []
    top_row_idx = None
    for idx in range(header_row_idx - 1, -1, -1):
        if idx < 0:
            break
        text_lower = _row_text_lower(rows_raw[idx][2])
        if "kepemilikan" in text_lower:
            top_row_idx = idx
            break
    if top_row_idx is not None:
        top_spans = rows_raw[top_row_idx][2]
        text_lower = _row_text_lower(top_spans)
        if "kepemilikan" in text_lower:
            # Tiap span: kolom mana saja yang overlap dengan bbox-nya
            span_col_ranges = []
            for s in top_spans:
                bbox = s.get("bbox") or (0, 0, 0, 0)
                cols = [j for j in range(num_cols) if _bbox_overlaps_col(bbox, j)]
                if cols:
                    span_col_ranges.append((" ".join((s.get("text") or "").split()), min(cols), max(cols)))
            span_col_ranges.sort(key=lambda x: x[1])
            # Gabung span yang berdekatan/overlap jadi satu sel
            merged = []
            for text, c0, c1 in span_col_ranges:
                if merged and c0 <= merged[-1][2] + 1:
                    merged[-1] = (merged[-1][0] + " " + text, merged[-1][1], max(merged[-1][2], c1))
                else:
                    merged.append((text.strip(), c0, c1))
            # Bangun list (text, colspan) urut per kolom: kosong atau satu sel span
            j = 0
            while j < num_cols:
                found = None
                for (text, c0, c1) in merged:
                    if c0 <= j <= c1:
                        found = (text, c0, c1)
                        break
                if found:
                    text, c0, c1 = found
                    header_top.append({"text": text, "colspan": c1 - c0 + 1})
                    j = c1 + 1
                else:
                    header_top.append({"text": "", "colspan": 1})
                    j += 1
        # Pastikan jumlah kolom header_top persis num_cols (kadang overlap bbox beda)
        total_span = sum(c.get("colspan", 1) for c in header_top)
        if total_span != num_cols and header_top:
            if total_span < num_cols:
                header_top.append({"text": "", "colspan": num_cols - total_span})
            else:
                # Kurangi colspan sel terakhir atau tambah sel kosong sampai pas
                while total_span > num_cols and header_top:
                    last = header_top[-1]
                    c = last.get("colspan", 1)
                    if c > 1:
                        last["colspan"] = c - 1
                        total_span -= 1
                    else:
                        header_top.pop()
                        total_span -= 1
                if total_span < num_cols:
                    header_top.append({"text": "", "colspan": num_cols - total_span})

    def _normalize_cell(s: str) -> str:
        """Trim dan satukan spasi/newline berlebih di isi sel."""
        if not s or not isinstance(s, str):
            return ""
        return " ".join(s.split())

    # Baris data: hanya baris di bawah header yang punya teks biru
    data_rows = []
    for idx in range(header_row_idx + 1, len(rows_raw)):
        _y, _page, row_spans = rows_raw[idx]
        if not any(s.get("is_blue") for s in row_spans):
            continue
        cells = [""] * num_cols
        for s in row_spans:
            if not s.get("is_blue"):
                continue
            bbox = s.get("bbox") or (0, 0, 0, 0)
            j = column_index_for_span(bbox)
            if 0 <= j < num_cols:
                text = (s.get("text") or "").strip()
                if not text:
                    continue
                if cells[j]:
                    cells[j] += " " + text
                else:
                    cells[j] = text
        # Normalisasi tiap sel: hilangkan spasi/newline berlebih
        cells = [_normalize_cell(c) for c in cells]
        data_rows.append(cells)

    # Perluas blok Kepemilikan yang hanya 2 kolom (Status+Perubahan) jadi 3 kolom
    # (Jumlah Saham, Saham Gabungan Per Investor, Persentase) agar template selalu benar.
    new_header_top = []
    col_offset = 0
    need_expand = False
    for h in header_top:
        c = h.get("colspan", 1)
        text = (h.get("text") or "").strip().lower()
        if "kepemilikan" in text and c == 2:
            new_header_top.append({"text": h.get("text", ""), "colspan": 3})
            col_offset += 3
            need_expand = True
        else:
            new_header_top.append(h)
            col_offset += c

    if need_expand:
        def _expand_row(row: list[str]) -> list[str]:
            out = []
            col_idx = 0
            for h in header_top:
                c = h.get("colspan", 1)
                text = (h.get("text") or "").strip().lower()
                if "kepemilikan" in text and c == 2:
                    out.append(row[col_idx] if col_idx < len(row) else "")
                    out.append("")  # kolom kosong untuk Saham Gabungan atau Persentase
                    out.append(row[col_idx + 1] if col_idx + 1 < len(row) else "")
                    col_idx += 2
                else:
                    for k in range(c):
                        out.append(row[col_idx + k] if col_idx + k < len(row) else "")
                    col_idx += c
            return out

        data_rows = [_expand_row(r) for r in data_rows]
        header_top = new_header_top
        num_cols = col_offset

    # Fallback 1: baris header utama punya sel berisi "Kepemilikan"
    if not header_top and header_cells:
        row_lower = " ".join(c.lower() for c in header_cells)
        if "kepemilikan" in row_lower:
            merged_top = []
            j = 0
            while j < num_cols:
                cell = (header_cells[j] if j < len(header_cells) else "") or ""
                if "kepemilikan" in cell.lower():
                    k = j
                    while k < num_cols and "kepemilikan" in ((header_cells[k] if k < len(header_cells) else "") or "").lower():
                        k += 1
                    merged_top.append({"text": " ".join(header_cells[j:k]) if k <= len(header_cells) else cell, "colspan": k - j})
                    j = k
                else:
                    merged_top.append({"text": "", "colspan": 1})
                    j += 1
            if sum(c.get("colspan", 1) for c in merged_top) == num_cols:
                header_top = merged_top

    # Fallback 2 (sintetis): deteksi pola 3 kolom (Jumlah Saham, Saham Gabungan, Persentase Kepemilikan)
    # lalu buat baris atas "Kepemilikan Per" dengan colspan 3 per grup - tidak bergantung ekstraksi PDF
    if not header_top and header_cells:
        def _cell(j):
            return ((header_cells[j] if j < len(header_cells) else "") or "").lower()

        synthetic = []
        j = 0
        group_num = 0
        while j < num_cols:
            if j + 3 <= num_cols:
                a, b, c = _cell(j), _cell(j + 1), _cell(j + 2)
                # Pola: kolom 1 = Jumlah Saham, 2 = Saham Gabungan..., 3 = Persentase Kepemilikan...
                like1 = "jumlah" in a and "saham" in a
                like2 = "gabungan" in b or ("saham" in b and "investor" in b)
                like3 = "persentase" in c or "kepemilikan" in c
                if like1 and (like2 or like3):
                    group_num += 1
                    synthetic.append({"text": "Kepemilikan Per" + (f" (periode {group_num})" if group_num > 1 else ""), "colspan": 3})
                    j += 3
                    continue
            synthetic.append({"text": "", "colspan": 1})
            j += 1
        if group_num > 0 and sum(c.get("colspan", 1) for c in synthetic) == num_cols:
            header_top = synthetic

    # Fallback 3: tabel punya "Status" dan "Perubahan" tapi sub-kolom Kepemilikan tidak terpisah
    # -> baris atas: kosong untuk kolom awal, "Kepemilikan Per" colspan 3 untuk 3 kolom sebelum "Perubahan"
    if not header_top and num_cols >= 10 and header_cells:
        row_lower = " ".join((c or "").lower() for c in header_cells)
        if "status" in row_lower and "perubahan" in row_lower and num_cols >= 4:
            header_top = [{"text": "", "colspan": 1} for _ in range(num_cols - 4)]
            header_top.append({"text": "Kepemilikan Per", "colspan": 3})
            header_top.append({"text": "", "colspan": 1})

    # Header dipakai template tetap; isi kolom datanya dari PDF (data_rows).
    # Jika ada header_top, bangun nama kolom dari strukturnya agar tiap blok Kepemilikan = 3 subkolom.
    if header_top and sum(c.get("colspan", 1) for c in header_top) == num_cols:
        template_header_row = build_template_header_row_from_header_top(header_top)
    else:
        template_header_row = build_template_header_row(num_cols)

    # Kembalikan format yang mendukung header 2 baris (header = template, data = dari PDF)
    if header_top and sum(c.get("colspan", 1) for c in header_top) == num_cols:
        return {
            "header_top": header_top,
            "header_row": template_header_row,
            "data": data_rows,
        }
    # Tanpa header atas: format flat (baris pertama = template header, sisanya data dari PDF)
    return template_header_row + data_rows


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
            table = [result["header_row"]] + result["data"]
            return jsonify({"table": table, "header_top": result["header_top"]})
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
