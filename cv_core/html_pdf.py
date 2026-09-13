"""
HTML/CSS based PDF renderer for the CV.

Why this exists:
  The old pipeline (python-docx -> LibreOffice --convert-to pdf) has two
  independent, server-only bugs (never happen in Microsoft Word on a PC):
    1) Arabic text-shaping bugs on some fonts/edits
    2) LibreOffice mis-placing inline images inside nested tables, causing
       them to overlap section headers
  This module renders the exact same CV data as HTML+CSS and converts it to
  PDF with WeasyPrint (a pure-Python PDF renderer). Every element gets an
  explicit position, so neither bug class can happen here.

  The .docx output is UNCHANGED and still produced by fill_cv() in app.py --
  this only replaces how the PDF is made.
"""

import base64
import shutil
import subprocess
import tempfile
from pathlib import Path

TEMPLATE_PATH = Path(__file__).parent / "cv_template.html"

ASSETS = Path(__file__).parent
LOGO_DIAMOND_ASAIL = ASSETS / "logo_diamond.jpeg"
LOGO_MHR = ASSETS / "logo_mhr.jpeg"
LOGO_DIAMOND_ALZAID = ASSETS / "logo_diamond_alzaid.jpeg"
LOGO_ALZAID_SMALL = ASSETS / "logo_alzaid_small.png"
LOGO_ALZAID_BIG = ASSETS / "logo_alzaid_big.png"


def _img_to_data_uri(path: Path) -> str:
    if not path.exists():
        return ""
    ext = path.suffix.lstrip(".").lower()
    mime = "jpeg" if ext in ("jpg", "jpeg") else ext
    b64 = base64.b64encode(path.read_bytes()).decode()
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


def _render_with_weasyprint(html: str, pdf_out: Path) -> bool:
    try:
        from weasyprint import HTML
    except Exception:
        return False
    try:
        HTML(string=html, base_url=str(TEMPLATE_PATH.parent)).write_pdf(str(pdf_out))
    except Exception:
        return False
    return pdf_out.exists() and pdf_out.stat().st_size > 0


def _render_with_wkhtmltopdf(html: str, pdf_out: Path) -> bool:
    wk = _find_wkhtmltopdf()
    if not wk:
        return False
    with tempfile.TemporaryDirectory() as td:
        html_path = Path(td) / "cv.html"
        html_path.write_text(html, encoding="utf-8")
        try:
            subprocess.run(
                [wk, "--enable-local-file-access", "--quiet",
                 str(html_path), str(pdf_out)],
                capture_output=True, timeout=60, check=False,
            )
        except Exception:
            return False
    return pdf_out.exists() and pdf_out.stat().st_size > 0


def render_pdf_via_html(data: dict, images: dict, pdf_out: Path,
                         image_sizes_in: dict = None,
                         agency: str = "Asail", experienced: bool = True) -> bool:
    """
    data:   same dict already used for the docx placeholders.
    images: same dict already used for the docx image placeholders
            (values are file paths, or None if not provided).
    image_sizes_in: optional {key: (width_in, height_in)} -- the exact
            dimensions already computed for the .docx version of each photo,
            so the PDF places photos at literally the same size.
    agency: "Asail" or "Al Zaid" -- selects the correct top-right logo,
            and (for Al Zaid) the extra page-2 logo block.
    experienced: True for the "Experienced" template (gold header title,
            "** EXPERIANCED **" label), False for "Non Experienced"
            (black header title, no label) -- matches the real templates.
    pdf_out: Path to write the final PDF to
    """
    image_sizes_in = image_sizes_in or {}
    html = TEMPLATE_PATH.read_text(encoding="utf-8")

    is_alzaid = "zaid" in agency.lower().replace(" ", "")

    logo_left = LOGO_DIAMOND_ALZAID if is_alzaid else LOGO_DIAMOND_ASAIL
    logo_right = LOGO_ALZAID_SMALL if is_alzaid else LOGO_MHR
    html = html.replace("{{LOGO_LEFT}}", _img_to_data_uri(logo_left))
    html = html.replace("{{LOGO_RIGHT}}", _img_to_data_uri(logo_right))

    html = html.replace("{{HEADER_COLOR}}", "#BF9000" if experienced else "#000000")
    html = html.replace("{{EXPERIENCED_LABEL}}", "** EXPERIANCED **" if experienced else "")

    if is_alzaid:
        big_uri = _img_to_data_uri(LOGO_ALZAID_BIG)
        page2_block = f'<div class="p2-logo"><img src="{big_uri}"></div>'
    else:
        page2_block = ""
    html = html.replace("{{PAGE2_LOGO_BLOCK}}", page2_block)

    # text placeholders
    for key, val in data.items():
        html = html.replace("{{%s}}" % key, "" if val is None else str(val))

    # image placeholders -> <img> tags, sized to match the .docx exactly
    for key, path in images.items():
        ph = "{{%s}}" % key
        if ph not in html:
            continue
        if path:
            uri = _img_to_data_uri(Path(path))
            if key in image_sizes_in:
                w_in, h_in = image_sizes_in[key]
                style = f'width:{w_in*25.4:.2f}mm; height:{h_in*25.4:.2f}mm; display:block;'
            else:
                style = 'width:100%; height:auto; display:block;'
            html = html.replace(ph, f'<img src="{uri}" style="{style}">')
        else:
            html = html.replace(ph, "")

    if _render_with_weasyprint(html, pdf_out):
        return True
    if _render_with_wkhtmltopdf(html, pdf_out):
        return True
    return False
