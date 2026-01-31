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
def is_blue_color(color) -> bool:
    """Cek apakah warna (sRGB int atau tuple float) dianggap biru."""
    if color is None:
        return False
    try:
        if isinstance(color, (int, float)) and not isinstance(color, bool):
            # sRGB integer 0xRRGGBB atau satu nilai (grayscale)
            if isinstance(color, float):
                return False
            r = (color >> 16) & 0xFF
            g = (color >> 8) & 0xFF
            b = color & 0xFF
        elif hasattr(color, "__len__") and len(color) >= 3:
            # Tuple/list (r, g, b) float 0–1
            r, g, b = float(color[0]) * 255, float(color[1]) * 255, float(color[2]) * 255
        else:
            return False
    except (TypeError, ValueError, IndexError):
        return False
    return b > r and b > g and b >= 80


def _as_list(val, default=None):
    """Pastikan nilai bisa di-iterate sebagai list."""
    if default is None:
        default = []
    if isinstance(val, list):
        return val
    if isinstance(val, (tuple, range)):
        return list(val)
    return default


def extract_blue_text_from_pdf(input_path: str) -> list[dict]:
    """Baca PDF, kembalikan list span yang warnanya biru (text, size, font, color)."""
    doc = fitz.open(input_path)
    blue_spans = []
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
                    color = span.get("color")
                    if is_blue_color(color):
                        raw_size = span.get("size", 12)
                        try:
                            size = float(raw_size) if raw_size is not None else 12
                        except (TypeError, ValueError):
                            size = 12
                        blue_spans.append({
                            "text": (span.get("text") or "").strip(),
                            "size": size,
                            "font": span.get("font", "helv"),
                            "page": page_num + 1,
                        })
    doc.close()
    return blue_spans


def create_pdf_with_blue_text(blue_spans: list[dict], output_path: str) -> None:
    """Buat PDF baru yang hanya berisi teks biru (tetap warna biru)."""
    doc = fitz.open()
    # Warna biru untuk teks (nilai 0–1 untuk PDF)
    blue_pdf = (0, 0, 1)
    line_height = 14
    margin = 50
    y = margin
    page_width = 595
    page_height = 842
    x_max = page_width - margin
    page = doc.new_page(width=page_width, height=page_height)
    for item in blue_spans:
        text = item["text"]
        if not text:
            continue
        try:
            size = min(float(item["size"]), 12)
        except (TypeError, ValueError):
            size = 12
        # Tambah label halaman sumber
        label = f"[hal {item['page']}] "
        full_text = label + text
        rect = fitz.Rect(margin, y, x_max, y + line_height * 2)
        # insert_textbox mengembalikan angka (sisa ruang), bukan teks overflow
        page.insert_textbox(rect, full_text, fontsize=size, color=blue_pdf)
        y += line_height
        if y > page_height - margin - line_height:
            page = doc.new_page(width=page_width, height=page_height)
            y = margin
    doc.save(output_path, garbage=4, deflate=True)
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
        buf = BytesIO()
        out_path = tempfile.mktemp(suffix=".pdf")
        try:
            create_pdf_with_blue_text(blue_spans, out_path)
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
    app.run(debug=True, port=5000)
