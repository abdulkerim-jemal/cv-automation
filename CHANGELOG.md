# Fixes applied — 2026-09-06

I tested every change by actually generating CVs from your app.py and
rendering them to PDF (via LibreOffice) so these are verified, not guesses.

## 1. PDF export was silently broken on non-Windows / no-Word setups
`docx2pdf` only works if Microsoft Word is installed and reachable via COM
(Windows/macOS + Word). Anywhere else, `made_pdf` just came back `False`
with no PDF, only the Word file.

**Fix:** `fill_cv()` now tries LibreOffice headless conversion
(`soffice --headless --convert-to pdf`) first — it works on Windows, macOS
and Linux with no Word dependency — and only falls back to `docx2pdf` if
LibreOffice isn't found. `Run_CV_App.bat` now also checks for LibreOffice
and tells you where to get it if missing.
→ Install LibreOffice (free): https://www.libreoffice.org/download/download/

## 2. Passport photo was tiny — a real bug, not a setting
`{{IMAGE_PASSPORT}}` sits in a plain paragraph in the Word templates (not a
table cell), so the sizing code's cell-lookup always failed and silently
fell back to a hardcoded `4.5 × 3.0 in` box — completely ignoring the
`PHOTO_BOX_IN` size that was supposed to apply. This is why the passport
image rendered small with a lot of wasted blank space around it (and an
entirely blank trailing page).

**Fix:** removed the bad fallback, and the passport box is now
`6.3 × 7.4 in` — verified it fills page 2 nicely with no overlap with the
Skills table, for both Asail and Al Zaid, Experienced and Non-Experienced.

## 3. Full-body photo made taller
The actual size was being driven by an internal "boost" factor applied
after fitting the photo to its box (not the box size itself, and not the
Word table's row height, which barely mattered — I traced this by
instrumenting the actual insertion code). Bumped that factor so the photo
renders noticeably larger while still fitting cleanly on page 1 for both
agencies (verified with photos of two different aspect ratios).

I also fixed a related correctness bug: the code was locking the
containing table row to the *pre-boost* (smaller) size instead of the
actual inserted picture size. LibreOffice is lenient and auto-expands
anyway, but real Microsoft Word can be stricter about "exact" row heights,
so this was a latent risk of clipping there even though it looked fine in
LibreOffice.

## 4. Faster to load
`easyocr` (which pulls in PyTorch) and `cv2` were imported unconditionally
at the top of the file, which is paid on every cold start of the app even
if you only ever use Tesseract. Both are now only imported the moment
you actually use them, and the EasyOCR reader is cached properly with
`st.cache_resource` so it's built once per app run, not on every rerun.

## 5. Simplified the crop screen
Removed:
- the right-hand "all regions" overlay preview panel, and
- the duplicate "Preview passport and OCR crops" expander

Both re-rendered a full-image composite on every single interaction with
no caching, and duplicated what the drag-crop canvas already shows. The
crop canvas itself is unchanged and now uses the full width.

## Known pre-existing cosmetic quirk (not something I changed)
On the Al Zaid template, the full-body photo's feet slightly overlap the
"Al Zaid Recruitment Office" watermark logo near the bottom of page 1.
I confirmed this already happened with the *original*, unmodified code too
— it's not something my changes introduced. If you'd like it cleaned up,
the simplest fix is nudging that logo image up slightly or shrinking it in
`Al Zaid/Experianced.docx` and `Al Zaid/Non_Experianced.docx` — happy to do
that too if you want it.

## What I didn't touch
Your OCR/MRZ/Gemini reading logic, the CV field layout, colors, and the
agency templates' content are all unchanged — only the sizing/PDF/perf/UI
issues above.
