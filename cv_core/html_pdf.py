"""
HTML/CSS based PDF renderer for the CV.

Why this exists:
  The old pipeline (python-docx -> LibreOffice --convert-to pdf) depends on
  LibreOffice correctly shaping Arabic text and substituting fonts it doesn't
  have installed. On the free Streamlit Cloud server this produced broken
  Arabic glyphs (e.g. "لا" rendering as a stray "U") even though the same
  file opened fine in Microsoft Word.

  This module renders the exact same CV data as HTML+CSS and converts it to
  PDF with wkhtmltopdf (a small, well-supported apt package). Browsers/
  WebKit shape Arabic correctly and predictably, so this sidesteps the whole
  class of font/shaping bugs.

  The .docx output is UNCHANGED and still produced by fill_cv() in app.py --
  this only replaces how the PDF is made.
"""

import base64
import shutil
import subprocess
import tempfile
from pathlib import Path

TEMPLATE_PATH = Path(__file__).parent / "cv_template.html"

# Put your two logo files here (same ones already used in the Word templates).
# Export them once from the .docx (word/media/imageN.jpeg) and drop them next
# to this file, e.g. cv_core/logo_diamond.jpeg and cv_core/logo_mhr.jpeg
LOGO_LEFT = Path(__file__).parent / "logo_diamond.jpeg"
LOGO_RIGHT = Path(__file__).parent / "logo_mhr.jpeg"


def _img_to_data_uri(path: Path) -> str:
    if not path.exists():
        return ""
    ext = path.suffix.lstrip(".").lower()
    mime = "jpeg" if ext in ("jpg", "jpeg") else ext
    b64 = base64.b64encode(path.read_bytes()).decode()
    return f"data:image/{mime};base64,{b64}"


def _img_file_to_data_uri_bytes(data: bytes, ext: str = "jpeg") -> str:
    mime = "jpeg" if ext.lower() in ("jpg", "jpeg") else ext.lower()
    b64 = base64.b64encode(data).decode()
    return f"data:image/{mime};base64,{b64}"


def _find_wkhtmltopdf():
    for name in ("wkhtmltopdf", "wkhtmltopdf.exe"):
        p = shutil.which(name)
        if p:
            return p
    for p in (r"C:\Program Files\wkhtmltopdf\bin\wkhtmltopdf.exe",
              r"C:\Program Files (x86)\wkhtmltopdf\bin\wkhtmltopdf.exe"):
        if Path(p).exists():
            return p
    return None


def render_pdf_via_html(data: dict, images: dict, pdf_out: Path) -> bool:
    """
    data:   same dict already used for the docx placeholders,
            e.g. {"NAME": "...", "PASSPORT_NO": "...", ...}
    images: same dict already used for the docx image placeholders,
            e.g. {"IMAGE_FULL": "/path/to/file.jpg", ...} (values are file
            paths, or None if not provided)
    pdf_out: Path to write the final PDF to
    """
    html = TEMPLATE_PATH.read_text(encoding="utf-8")

    # logos
    html = html.replace("{{LOGO_LEFT}}", _img_to_data_uri(LOGO_RIGHT))   # diamond logo, left side in header
    html = html.replace("{{LOGO_RIGHT}}", _img_to_data_uri(LOGO_LEFT))   # mhr logo, right side in header

    # text placeholders
    for key, val in data.items():
        html = html.replace("{{%s}}" % key, "" if val is None else str(val))

    # image placeholders -> <img> tags
    for key, path in images.items():
        ph = "{{%s}}" % key
        if ph not in html:
            continue
        if path:
            uri = _img_to_data_uri(Path(path))
            html = html.replace(ph, f'<img src="{uri}">')
        else:
            html = html.replace(ph, "")

    with tempfile.TemporaryDirectory() as td:
        html_path = Path(td) / "cv.html"
        html_path.write_text(html, encoding="utf-8")

        wk = _find_wkhtmltopdf()
        if not wk:
            return False
        try:
            subprocess.run(
                [wk, "--enable-local-file-access", "--quiet",
                 str(html_path), str(pdf_out)],
                capture_output=True, timeout=60, check=False,
            )
        except Exception:
            return False

    return pdf_out.exists() and pdf_out.stat().st_size > 0
