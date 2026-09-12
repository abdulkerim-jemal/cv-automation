# CV Automation — Setup Guide

## What's new (this update)

- **PDF export actually works now, on any OS** – it previously only worked if
  Microsoft Word was installed (via `docx2pdf`), so on machines without Word
  (e.g. most non-Windows setups) PDF generation silently failed. The app now
  converts with **LibreOffice** first (free, headless, works on Windows/macOS/
  Linux) and only falls back to Word/`docx2pdf` if LibreOffice isn't found.
  Install LibreOffice from https://www.libreoffice.org/download/download/ if
  you don't already have it or Word - the Word (.docx) file is always
  produced either way, PDF is the only part that needs one of the two.
- **Passport photo bug fixed** – it was silently falling back to a small,
  hardcoded box instead of the intended size. It's now much bigger and fills
  its own page cleanly.
- **Full‑body photo made taller** without breaking the page layout.
- **Faster to load** – EasyOCR (and OpenCV) are no longer imported until you
  actually use them, so the app starts noticeably faster if you're just
  using Tesseract.
- **Simplified crop screen** – removed the extra "all regions" preview panel
  and the duplicate crop-preview box; the drag‑crop canvas already shows the
  selected region directly on the photo.

## Previous update

- **Drag‑crop** – Use the canvas to draw rectangles for each region (passport, 3×4, full body, and OCR reading zone).
- **Separate OCR crop** – Define a custom crop area for MRZ/text reading to improve accuracy.
- **EasyOCR support** – Optionally use EasyOCR (more accurate than Tesseract) for printed text extraction.

## Current version

This build uses a clean three‑step dashboard UI (crop screen no longer has a
side preview panel):

1. **Upload** – choose agency and experience, then upload the composite photo.
2. **Crop** – use either number boxes or the **drag‑crop canvas** to adjust the passport, 3×4, full‑body, and (optionally) OCR reading zone.
3. **Details** – use the **🚀 Smart Read** button (OCR + MRZ + Gemini) and review the extracted information.
4. **Export** – generated files are saved loose inside dated output folders.

## Passport reading

The Smart Read flow:

- Tesseract (or EasyOCR) reads the passport and MRZ first.
- When the MRZ is successfully parsed, its Name, Passport No., Date of Birth and Date of Expiry are preferred because the MRZ has checksum validation.
- **Place of Birth** and **Issue Date** are read from the printed passport (Gemini helps if an API key is supplied).
- If a **separate OCR crop** is enabled, that region is used for text extraction, while the full passport is still used for Gemini (if available).
- **Date of Issue** is normally inferred from the expiry date using the selected **5‑year** or **10‑year** validity. If expiry cannot be read, it falls back to visual OCR.
- Dates are normalized to `DD/MM/YYYY` before being placed into the dropdowns.

## Date ranges

The dropdown year ranges are hard‑limited to:

- Date of Birth: **1981–2007**
- Date of Issue: **2019–2030**
- Date of Expiry: **2021–2038**

## Smart marital/children defaults

After the DOB is known, the form suggests:

- Age ≤ 25 → SINGLE / NIL
- Age 26–28 → MARRIED / 1
- Age 29–31 → MARRIED / 2
- Age 32–35 → MARRIED / 3
- Age > 35 → MARRIED / 4

Both dropdowns remain editable.

## Output structure

The app creates today's root folder automatically:

```text
Output_YYYY-MM-DD/
├── Asail/
│   ├── candidate.docx
│   ├── candidate.pdf
│   ├── candidate_passport.jpg
│   ├── candidate_3x4.jpg
│   └── candidate_full.jpg
└── Al Zaid/
    └── ...