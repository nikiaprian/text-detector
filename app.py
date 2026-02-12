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
    return {
        "text": (span.get("text") or "").strip(),
        "size": size,
        "font": span.get("font", "helv"),
        "page": page_num + 1,
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


@app.route("/")
def index():
    return render_template("index.html")


@app.route("/api/extract-blue", methods=["POST"])
def extract_blue():
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
        blue_spans = extract_blue_text_from_pdf(tmp_in_path)
        if not blue_spans:
            return jsonify({
                "error": "Tidak ada teks warna biru ditemukan di PDF ini.",
                "hint": "Pastikan teks benar-benar menggunakan warna biru (bukan hitam/abu-abu)."
            }), 422
        output_style = (request.form.get("output_style") or "paragraph").strip().lower()
        if output_style not in OUTPUT_STYLES:
            output_style = "paragraph"

        buf = BytesIO()
        out_path = tempfile.mktemp(suffix=".pdf")
        try:
            create_pdf_with_blue_text(blue_spans, out_path, output_style=output_style)
            with open(out_path, "rb") as f:
                buf.write(f.read())
        finally:
            if os.path.exists(out_path):
                try:
                    os.unlink(out_path)
                except Exception:
                    pass
        buf.seek(0)
        return send_file(
            buf,
            mimetype="application/pdf",
            as_attachment=True,
            download_name="teks_biru.pdf",
        )
    except Exception as e:
        return jsonify({"error": str(e)}), 500
    finally:
        if tmp_in_path and os.path.exists(tmp_in_path):
            try:
                os.unlink(tmp_in_path)
            except Exception:
                pass


if __name__ == "__main__":
    app.run(host="0.0.0.0", port=5000, debug=True)
