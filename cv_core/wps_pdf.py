"""
Convert the *actual* generated .docx to PDF by automating WPS Office itself
(Windows only), instead of re-rendering the CV independently in HTML/CSS.

Why this exists:
  html_pdf.py rebuilds the CV from scratch as HTML and rasterizes that with
  WeasyPrint. That's a totally different layout engine from Word/WPS, so the
  PDF can never be pixel-identical to the .docx -- different font metrics,
  line-breaking, image placement math, etc.

  WPS's own "switch extension" / Save As PDF works perfectly because it's
  the SAME engine that renders the .docx on screen -- there's no
  re-implementation step, so nothing can drift.

  This module reproduces exactly that: it opens the real .docx you already
  generated in WPS Writer (via COM automation) and asks WPS to export it,
  so the PDF is byte-for-byte the same layout you'd get doing it by hand.

Requirements:
  - Windows
  - WPS Office installed (Kingsoft WPS Writer)
  - pywin32 (pip install pywin32)

Only used when running locally on the machine that has WPS installed --
this will simply return False (and let the caller fall back to the next
option) on any other OS or if WPS/pywin32 isn't available.
"""

import os
from pathlib import Path

# wdFormatPDF -- same constant Word uses; WPS mirrors it for compatibility.
WPS_PDF_FORMAT = 17

# Try newer unified ProgID first, then the classic Writer-specific one.
WPS_PROGIDS = ("wps.Application", "kwps.Application")


def convert_docx_to_pdf_via_wps(docx_path, pdf_path, timeout_s: int = 90) -> bool:
    """
    Open docx_path in WPS Writer and export it to pdf_path.
    Returns True on success, False if WPS/pywin32 isn't available or the
    conversion fails for any reason (caller should fall back to another
    method in that case -- this never raises).
    """
    if os.name != "nt":
        return False

    docx_path = Path(docx_path).resolve()
    pdf_path = Path(pdf_path).resolve()
    if not docx_path.exists():
        return False

    try:
        import pythoncom
        import win32com.client as win32
    except Exception:
        return False

    pythoncom.CoInitialize()
    app = None
    doc = None
    try:
        for progid in WPS_PROGIDS:
            try:
                app = win32.gencache.EnsureDispatch(progid)
                break
            except Exception:
                app = None
                continue
        if app is None:
            return False

        app.Visible = False
        try:
            app.DisplayAlerts = False
        except Exception:
            pass

        doc = app.Documents.Open(str(docx_path))

        # Prefer ExportAsFixedFormat when present (more faithful PDF export,
        # same call WPS makes internally for "Export to PDF"); fall back to
        # a plain SaveAs2 with the PDF file format constant.
        try:
            doc.ExportAsFixedFormat(str(pdf_path), WPS_PDF_FORMAT)
        except Exception:
            doc.SaveAs2(str(pdf_path), FileFormat=WPS_PDF_FORMAT)

        return pdf_path.exists() and pdf_path.stat().st_size > 0
    except Exception:
        return False
    finally:
        try:
            if doc is not None:
                doc.Close(False)
        except Exception:
            pass
        try:
            if app is not None:
                app.Quit()
        except Exception:
            pass
        pythoncom.CoUninitialize()
