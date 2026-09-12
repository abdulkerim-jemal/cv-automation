from pathlib import Path
from urllib.request import urlopen

ROOT=Path(__file__).resolve().parent
OUT=ROOT/'assets'/'tessdata'; OUT.mkdir(parents=True,exist_ok=True)
url='https://github.com/tesseract-ocr/tessdata_fast/raw/main/amh.traineddata'
out=OUT/'amh.traineddata'
with urlopen(url,timeout=60) as r: out.write_bytes(r.read())
print('Downloaded',out)
print('Set TESSDATA_PREFIX to',OUT.parent)
