from __future__ import annotations

from pathlib import Path
import fitz

from .ocr import tesseract_amharic


def ocr_pdf_to_searchable(input_pdf, output_pdf, dpi=220, postprocess=None):
    """Render each PDF page, OCR it, and add invisible text over the original.
    The original page is retained, so visual layout is not changed."""
    src=fitz.open(str(input_pdf)); out=fitz.open()
    try:
        for page in src:
            pix=page.get_pixmap(dpi=dpi,alpha=False)
            img=__import__('PIL.Image',fromlist=['Image']).Image.frombytes("RGB",[pix.width,pix.height],pix.samples)
            text,engine=tesseract_amharic(img,psm=6)
            if postprocess:
                text=postprocess(text)
            new=out.new_page(width=page.rect.width,height=page.rect.height)
            new.show_pdf_page(new.rect,src,page.number)
            # Approximate line boxes using Tesseract TSV when available; for a
            # lightweight searchable PDF, place the text as transparent page text.
            if text.strip():
                scale_x=page.rect.width/img.width; scale_y=page.rect.height/img.height
                y=10
                for line in text.splitlines():
                    if line.strip():
                        new.insert_text((4,y),line,fontsize=max(5,7*scale_x),fontname="helv",render_mode=3)
                        y += 8
        out.save(str(output_pdf),garbage=4,deflate=True)
    finally:
        src.close(); out.close()
    return Path(output_pdf)
