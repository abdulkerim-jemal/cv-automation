# CV Automation — repaired build

## What changed
- Local OCR is now the default; Gemini is optional and never required for passport reading.
- Tesseract is checked as an **executable**, not merely as a `pytesseract` import.
- MRZ uses PassportEye when installed, with a custom TD3 candidate/threshold/PSM/checksum fallback.
- Passport fields are cross-checked and uncertain/missing values are shown for manual review.
- Images are crop-to-fill fixed frames; source aspect ratio is preserved and source dimensions cannot resize Word tables.
- DOCX/PDF output is fixed-layout and validated to exactly two pages.
- PDF output is generated first from the original agency template pagination, then fixed-frame images are overlaid; the Word file is a print-locked two-page representation of the final PDF so Word cannot reflow the CV.
- Desktop output is `Desktop/output_YYYY-MM-DD/Asail/FullName/` or `Desktop/output_YYYY-MM-DD/AlZaid/FullName/` with `passport.jpg`, `3x4.jpg`, `full.jpg`, `CV.docx`, and `CV.pdf`.
- UI has a tight header, crop controls on the right, candidate details split left/right, OCR diagnostics, and live preview.
- Optional Amharic OCR supports Tesseract `amh`, a custom PyTorch CNN+BiLSTM+CTC architecture, Gemini text post-processing with key rotation, and searchable-PDF processing.

## Installation
1. Install Python 3.11+.
2. Install Tesseract OCR 5.x and ensure the executable is on PATH, or set `TESSERACT_CMD`.
3. Install LibreOffice and ensure `soffice` is on PATH.
4. `python -m pip install -r requirements.txt`
5. Optional MRZ/EasyOCR/CRNN features: `python -m pip install -r requirements-amharic.txt`
6. Run `streamlit run app.py` or double-click `Run_CV_App.bat` on Windows.

### Amharic language data
Run `python download_tessdata.py` to download the official `amh.traineddata` file. The project also includes a custom CRNN architecture and training function; a genuinely fine-tuned model cannot be supplied without a labeled Amharic training dataset, so training remains an optional user step.

## OCR behaviour
- Tesseract package OK is not treated as OCR availability unless the executable responds.
- PassportEye is attempted first when installed.
- Fallback MRZ processing searches multiple bottom-region candidates and preprocessing variants and validates TD3 checksums.
- MRZ confidence is exposed as a 0–100 score.
- Gemini errors, missing keys, quota errors, and timeouts fall back to local OCR.

## Gemini key rotation
Enter one key per line in the UI or set `GEMINI_API_KEYS` as a comma/newline-separated environment variable. Keys are tried sequentially on rate-limit/quota/transient failures. Do not commit keys to this project.

## Tests
Run:

```bash
pytest -q tests
```

The automated tests cover date normalization, MRZ checksum logic, fixed-frame dimensions, and template presence. The rendering test matrix used during this repair generated all four agency/experience combinations at exactly two PDF pages, including very wide/tall source images.

## Remaining limitations
- PassportEye and EasyOCR are optional because their native dependencies vary by Windows installation.
- Amharic accuracy depends on the installed Tesseract `amh` data or a model trained on an appropriate labeled dataset.
- The final DOCX is intentionally print-locked as page images. This is deliberate: editable Word reflow was the direct source of the original extra-page/image-distortion failures.
