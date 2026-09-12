from __future__ import annotations

import io
import json
import os
import re
import shutil
import subprocess
import tempfile
import urllib.error
import urllib.request
from dataclasses import dataclass, asdict
from datetime import datetime
from pathlib import Path
from typing import Optional

import numpy as np
from PIL import Image, ImageFilter, ImageOps


@dataclass
class EngineStatus:
    pytesseract_package: bool
    tesseract_executable: bool
    tesseract_path: str
    tesseract_version: str
    languages: list[str]
    passporteye: bool
    easyocr: bool
    mrz_package: bool


@dataclass
class MRZResult:
    lines: list[str]
    valid: bool
    valid_score: int
    valid_number: bool
    valid_dob: bool
    valid_expiry: bool
    valid_composite: bool
    name: str = ""
    passport_no: str = ""
    nationality: str = ""
    dob: str = ""
    expiry: str = ""
    sex: str = ""
    method: str = ""
    raw_text: str = ""

    def to_dict(self):
        return asdict(self)


MONTHS = {m:i for i,m in enumerate([
    "JAN","FEB","MAR","APR","MAY","JUN","JUL","AUG","SEP","OCT","NOV","DEC"
], 1)}
MRZ_FIX = str.maketrans({"O":"0","Q":"0","D":"0","U":"0","I":"1","L":"1","Z":"2","A":"4","S":"5","G":"6","T":"7","B":"8"})


def find_tesseract() -> tuple[str|None, str]:
    candidates=[]
    env=os.environ.get("TESSERACT_CMD")
    if env: candidates.append(env)
    candidates += [shutil.which("tesseract"),
        r"C:\Program Files\Tesseract-OCR\tesseract.exe",
        r"C:\Program Files (x86)\Tesseract-OCR\tesseract.exe"]
    for c in candidates:
        if c and Path(c).exists():
            try:
                p=subprocess.run([c,"--version"],capture_output=True,text=True,timeout=8)
                return c, (p.stdout or p.stderr).splitlines()[0] if p.returncode==0 else ""
            except Exception:
                continue
    return None, ""


def get_engine_status():
    try:
        import pytesseract
        py_ok=True
    except Exception:
        py_ok=False
    cmd,version=find_tesseract()
    langs=[]
    if cmd:
        try:
            p=subprocess.run([cmd,"--list-langs"],capture_output=True,text=True,timeout=10)
            langs=[x.strip() for x in (p.stdout or "").splitlines()[1:] if x.strip()]
        except Exception: pass
    try: import passporteye; pe=True
    except Exception: pe=False
    try: import easyocr; eo=True
    except Exception: eo=False
    try: from mrz.checker.td3 import TD3CodeChecker; mp=True
    except Exception: mp=False
    return EngineStatus(py_ok,bool(cmd),cmd or "",version,langs,pe,eo,mp)


def configure_tesseract():
    try:
        import pytesseract
    except Exception:
        return None
    cmd,_=find_tesseract()
    if cmd:
        pytesseract.pytesseract.tesseract_cmd=cmd
    return pytesseract if cmd else None


def preprocess_for_ocr(img: Image.Image, min_width=1800):
    gray=img.convert("L")
    if gray.width < min_width:
        scale=min_width/max(1,gray.width)
        gray=gray.resize((int(gray.width*scale),int(gray.height*scale)),Image.Resampling.LANCZOS)
    gray=ImageOps.autocontrast(gray,cutoff=1).filter(ImageFilter.SHARPEN)
    return gray


def ocr_variants(img):
    gray=img.convert("L")
    yield "gray",gray
    yield "contrast",ImageOps.autocontrast(gray,cutoff=1)
    arr=np.asarray(gray)
    # Otsu without OpenCV dependency.
    hist=np.bincount(arr.ravel(),minlength=256).astype(float)
    total=hist.sum(); sum_total=np.dot(np.arange(256),hist)
    sum_b=wb=0.0; best_t=127; best_var=0.0
    for t in range(256):
        wb += hist[t]
        if wb==0: continue
        wf=total-wb
        if wf==0: break
        sum_b += t*hist[t]
        mb=sum_b/wb; mf=(sum_total-sum_b)/wf
        v=wb*wf*(mb-mf)**2
        if v>best_var: best_var=v; best_t=t
    yield "otsu",gray.point(lambda p:255 if p>best_t else 0)
    try:
        import cv2
        ad=cv2.adaptiveThreshold(arr,255,cv2.ADAPTIVE_THRESH_GAUSSIAN_C,cv2.THRESH_BINARY,31,11)
        yield "adaptive",Image.fromarray(ad)
        clahe=cv2.createCLAHE(clipLimit=2.0,tileGridSize=(8,8)).apply(arr)
        yield "clahe",Image.fromarray(clahe)
    except Exception: pass


def _clean_line(line):
    return re.sub(r"[^A-Z0-9<]", "", line.upper())


def _candidate_lines(text):
    raw=[_clean_line(x) for x in text.splitlines()]
    raw=[x for x in raw if x]
    candidates=[]
    for i,a in enumerate(raw):
        for b in raw[i+1:i+4]:
            if 36<=len(a)<=48 and 36<=len(b)<=48:
                aa=(a+"<"*44)[:44]; bb=(b+"<"*44)[:44]
                if aa.startswith("P<") or (aa[0:2] in ("P<","V<") and len(bb)>=40):
                    candidates.append((aa,bb))
    return candidates


def _mrz_char_value(c):
    if c.isdigit(): return int(c)
    if c=='<': return 0
    return ord(c)-ord('A')+10 if 'A'<=c<='Z' else 0


def mrz_checksum(s):
    weights=(7,3,1)
    return sum(_mrz_char_value(c)*weights[i%3] for i,c in enumerate(s))%10


def _check_field(line,start,end,check_idx):
    if len(line)<check_idx+1: return False
    raw=line[start:end]; d=line[check_idx]
    return d.isdigit() and mrz_checksum(raw)==int(d)


def parse_td3(line1,line2,method="custom",raw_text=""):
    line1=(line1+"<"*44)[:44]; line2=(line2+"<"*44)[:44]
    valid_number=_check_field(line2,0,9,9)
    valid_dob=_check_field(line2,13,19,19)
    valid_exp=_check_field(line2,21,27,27)
    composite_raw=line2[0:10]+line2[13:20]+line2[21:28]+line2[28:43]
    valid_composite=len(line2)>=44 and line2[43].isdigit() and mrz_checksum(composite_raw)==int(line2[43])
    length_ok=(len(line1)==44 and len(line2)==44)
    score=0
    score += 25 if valid_number else 0
    score += 20 if valid_dob else 0
    score += 20 if valid_exp else 0
    score += 25 if valid_composite else 0
    score += 5 if length_ok else 0
    score += 5 if line1.startswith("P<") else 0
    surname_names=line1[5:44].rstrip("<").split("<<",1)
    surname=surname_names[0].replace("<"," ").strip() if surname_names else ""
    given=surname_names[1].replace("<"," ").strip() if len(surname_names)>1 else ""
    name=re.sub(r"\s+"," ",f"{given} {surname}").strip()
    number=line2[:9].replace("<","")
    dob=_decode_mrz_date(line2[13:19],True) if valid_dob or line2[13:19].translate(MRZ_FIX).isdigit() else ""
    expiry=_decode_mrz_date(line2[21:27],False) if valid_exp or line2[21:27].translate(MRZ_FIX).isdigit() else ""
    return MRZResult([line1,line2],score>=90,score,valid_number,valid_dob,valid_exp,valid_composite,name.upper(),number.upper(),line2[10:13],dob,expiry,line2[20],method,raw_text)


def _decode_mrz_date(raw,is_birth):
    s=raw.translate(MRZ_FIX)
    if len(s)!=6 or not s.isdigit(): return ""
    yy,mm,dd=int(s[:2]),int(s[2:4]),int(s[4:6])
    if not (1<=mm<=12 and 1<=dd<=31): return ""
    current=datetime.today().year%100
    year=(1900+yy) if is_birth and yy>current else (2000+yy)
    try: datetime(year,mm,dd)
    except ValueError: return ""
    return f"{dd:02d}/{mm:02d}/{year:04d}"


def passporteye_read(img: Image.Image):
    try:
        from passporteye import read_mrz
    except Exception:
        return None
    tmp=None
    try:
        f=tempfile.NamedTemporaryFile(suffix=".jpg",delete=False)
        tmp=f.name; f.close(); img.convert("RGB").save(tmp,"JPEG",quality=95)
        result=read_mrz(tmp)
        if result is None: return None
        d=result.to_dict() if hasattr(result,"to_dict") else {}
        # PassportEye exposes fields as attributes in common releases.
        def g(*names):
            for n in names:
                v=d.get(n) if isinstance(d,dict) else getattr(result,n,None)
                if v not in (None,""): return str(v)
            return ""
        line1=g("raw_text")
        raw=g("aux")
        lines=[]
        if isinstance(d,dict) and d.get("raw_text"): lines=[x for x in str(d["raw_text"]).splitlines() if x]
        if len(lines)>=2:
            parsed=parse_td3(lines[-2],lines[-1],"PassportEye",str(raw))
        else:
            parsed=MRZResult([],bool(getattr(result,"valid",False)),int(getattr(result,"valid_score",0) or 0),
                bool(getattr(result,"valid_number",False)),bool(getattr(result,"valid_date_of_birth",False)),
                bool(getattr(result,"valid_expiration_date",False)),bool(getattr(result,"valid_composite",False)),
                g("names"),g("number"),g("nationality"),_decode_mrz_date(g("date_of_birth"),True),
                _decode_mrz_date(g("expiration_date"),False),g("sex"),"PassportEye",line1)
        return parsed
    except Exception:
        return None
    finally:
        if tmp:
            try: os.unlink(tmp)
            except OSError: pass


def custom_mrz_read(img: Image.Image):
    """Fast MRZ reader.

    The previous implementation tried 7 crop fractions × 5 image variants ×
    5 Tesseract PSM modes (175 OCR processes). That made Smart Read feel
    frozen. We now use the OCR text already produced by the normal reader and
    only one small dedicated MRZ pass when necessary.
    """
    pyt = configure_tesseract()
    if pyt is None:
        return None, []

    prep = preprocess_for_ocr(img, min_width=1400)
    w, h = prep.size
    zone = prep.crop((0, int(h * 0.68), w, h))

    diagnostics = []
    best = None

    try:
        text = pyt.image_to_string(
            zone,
            lang="eng",
            config="--oem 3 --psm 6 -c tessedit_char_whitelist=ABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789<",
            timeout=8,
        )
    except Exception as e:
        return None, [{"error": str(e)}]

    diagnostics.append({"method": "fast-mrz", "text": text})
    candidates = _candidate_lines(text)
    for a, b in candidates:
        r = parse_td3(a, b, "Tesseract fast", text)
        if best is None or r.valid_score > best.valid_score:
            best = r
        if r.valid_score >= 90:
            break

    return best, diagnostics



def normalize_date_token(token):
    token=str(token or "").strip().upper().replace(","," ")
    token=re.sub(r"\s+"," ",token)
    m=re.match(r"^(\d{4})[\s/\-.]+(\d{1,2})[\s/\-.]+(\d{1,2})$",token)
    if m:
        y,mo,d=map(int,m.groups())
        try: datetime(y,mo,d); return f"{d:02d}/{mo:02d}/{y:04d}"
        except ValueError:return None
    m=re.match(r"^(\d{1,2})[\s/\-.]+(\d{1,2})[\s/\-.]+(\d{2,4})$",token)
    if m:
        d,mo,y=m.groups(); y=int(y); y=y if y>=1000 else (2000+y if y<50 else 1900+y)
        try: datetime(y,int(mo),int(d)); return f"{int(d):02d}/{int(mo):02d}/{y:04d}"
        except ValueError:return None
    m=re.match(r"^(\d{1,2})[\s/\-.]+([A-Z]+)[\s/\-.]+(\d{2,4})$",token)
    if m:
        d,mon,y=m.groups(); mo=MONTHS.get(mon,MONTHS.get(mon[:3])); y=int(y); y=y if y>=1000 else (2000+y if y<50 else 1900+y)
        if mo:
            try: datetime(y,mo,int(d)); return f"{int(d):02d}/{mo:02d}/{y:04d}"
            except ValueError: pass
    return None


def extract_labeled_fields(text):
    lines=[re.sub(r"\s+"," ",x).strip(" :|-") for x in (text or "").upper().replace("\r","\n").splitlines()]
    lines=[x for x in lines if x]
    out={"PLACE_OF_BIRTH":"","ISSUE_DATE":"","NATIONALITY":""}
    def next_value(patterns,max_ahead=3):
        for i,line in enumerate(lines):
            if any(re.search(p,line) for p in patterns):
                for p in patterns:
                    m=re.search(p+r"\s*[:\-]?\s*(.+)$",line)
                    if m and m.group(1).strip(): return m.group(1).strip()
                for j in range(i+1,min(len(lines),i+1+max_ahead)):
                    c=lines[j].strip(" :|-.")
                    if c and not re.search(r"DATE\s+OF|NATIONALITY|SEX|PASSPORT|SURNAME|GIVEN|NAME|PLACE\s+OF",c): return c
        return ""
    pob=next_value([r"PLACE\s+OF\s+BIRTH",r"PLACE\s*BIRTH"],4)
    if pob: out["PLACE_OF_BIRTH"]=re.sub(r"\b(DATE|SEX|NATIONALITY|PASSPORT)\b.*$","",pob).strip(" :|-")
    issue=next_value([r"DATE\s+OF\s+ISSUE",r"ISSUE\s+DATE",r"DATE\s+ISSUED"],3)
    dates=re.findall(r"\b(?:\d{1,2}[\s./-]+(?:\d{1,2}|[A-Z]{3,9})[\s./-]+\d{2,4}|\d{4}[\s./-]+\d{1,2}[\s./-]+\d{1,2})\b",issue)
    if dates: out["ISSUE_DATE"]=normalize_date_token(dates[0]) or ""
    nat=next_value([r"NATIONALIT[YV]"],2)
    if nat: out["NATIONALITY"]=re.split(r"\s{2,}|DATE|SEX|PLACE",nat)[0].strip()
    return out


def general_ocr(img, use_easyocr=False):
    """Fast passport text OCR.

    One PSM-6 pass normally reads the Ethiopian passport fields very well.
    A PSM-11 pass is used as a second, quick recovery pass; both outputs are
    returned together so the parser can choose the best date candidates.
    """
    if use_easyocr:
        try:
            import easyocr
            reader = get_easyocr_reader()
            text = " ".join(
                reader.readtext(np.asarray(img.convert("RGB")), detail=0, paragraph=True)
            )
            if text.strip():
                return text, "EasyOCR"
        except Exception:
            pass

    pyt = configure_tesseract()
    if pyt is None:
        return "", "Unavailable"

    prep = preprocess_for_ocr(img, min_width=1400)
    texts = []
    for psm in (6, 11):
        try:
            value = pyt.image_to_string(
                prep,
                lang="eng",
                config=f"--oem 3 --psm {psm}",
                timeout=8,
            )
            if value.strip():
                texts.append(value)
        except Exception:
            pass

    return "\n".join(texts), "Tesseract 5.x"




try:
    import streamlit as _st
except Exception:
    _st=None

if _st:
    @_st.cache_resource(show_spinner="Loading EasyOCR model (first time only)...")
    def get_easyocr_reader():
        import easyocr
        return easyocr.Reader(["en"],gpu=False)
else:
    def get_easyocr_reader():
        import easyocr
        return easyocr.Reader(["en"],gpu=False)


def _age_from_dob(dob):
    try:
        d=datetime.strptime(dob,"%d/%m/%Y"); t=datetime.today()
        return str(t.year-d.year-((t.month,t.day)<(d.month,d.day)))
    except Exception:return ""


def read_passport_local(
    image: Image.Image,
    ocr_image: Image.Image | None = None,
    use_easyocr=False,
    use_passporteye=False,
):
    """Read passport fields quickly with Tesseract/EasyOCR.

    PassportEye remains available but is opt-in because its extra processing is
    unnecessary for ordinary Ethiopian passport OCR and can make the UI slow.
    """
    target = ocr_image or image

    full_text, engine = general_ocr(target, use_easyocr)

    # First try MRZ candidates already present in the normal OCR output.
    mrz = None
    candidates = _candidate_lines(full_text)
    for a, b in candidates:
        r = parse_td3(a, b, "Tesseract OCR text", full_text)
        if mrz is None or r.valid_score > mrz.valid_score:
            mrz = r
        if r.valid_score >= 90:
            break

    # One dedicated MRZ pass only if the normal OCR did not produce a usable MRZ.
    diagnostics = []
    if mrz is None or mrz.valid_score < 90:
        mrz2, diagnostics = custom_mrz_read(target)
        if mrz2 and (mrz is None or mrz2.valid_score > mrz.valid_score):
            mrz = mrz2

    # PassportEye is available for troubleshooting/deep MRZ checks, but not on
    # the default fast path.
    if use_passporteye:
        try:
            pe = passporteye_read(target)
            if pe and (mrz is None or pe.valid_score > mrz.valid_score):
                mrz = pe
        except Exception:
            pass

    labeled = extract_labeled_fields(full_text)
    info = {
        "NAME": "",
        "PASSPORT_NO": "",
        "DOB": "",
        "EXPIRY_DATE": "",
        "ISSUE_DATE": labeled["ISSUE_DATE"],
        "HOME_ADDRESS": labeled["PLACE_OF_BIRTH"],
        "AGE": "",
        "_warning": "",
        "_barcode_raw": "",
        "_mrz": mrz.to_dict() if mrz else {},
        "_diagnostics": diagnostics,
        "_ocr_text": full_text,
        "_engine": engine,
    }

    if mrz:
        info.update({
            "NAME": mrz.name,
            "PASSPORT_NO": mrz.passport_no,
            "DOB": mrz.dob,
            "EXPIRY_DATE": mrz.expiry,
        })
        info["AGE"] = _age_from_dob(info["DOB"])

    warnings = []
    if not mrz:
        warnings.append("MRZ not detected; printed passport fields are still read.")
    elif mrz.valid_score < 90:
        warnings.append(
            f"MRZ confidence {mrz.valid_score}/100; verify the fields visually."
        )
    if engine == "Unavailable":
        warnings.append(
            "Tesseract executable was not detected. Install Tesseract OCR or set TESSERACT_CMD."
        )
    for k in ("NAME", "PASSPORT_NO", "DOB", "EXPIRY_DATE", "ISSUE_DATE", "HOME_ADDRESS"):
        if not info[k]:
            warnings.append(f"Still missing: {k}.")
    info["_warning"] = " ".join(dict.fromkeys(warnings)) or "Local OCR complete."
    return info



def gemini_keys_from_text(value):
    return [x.strip() for x in re.split(r"[\n,;]+",value or "") if x.strip()]


def gemini_json_with_rotation(image: Image.Image, keys, model, prompt, schema=None, timeout=25):
    buf=io.BytesIO(); image.convert("RGB").save(buf,"JPEG",quality=92)
    b64=__import__("base64").b64encode(buf.getvalue()).decode("ascii")
    keys=list(keys)
    errors=[]
    for idx,key in enumerate(keys):
        payload={"contents":[{"parts":[{"text":prompt},{"inline_data":{"mime_type":"image/jpeg","data":b64}}]},],"generationConfig":{"temperature":0,"response_mime_type":"application/json"}}
        if schema: payload["generationConfig"]["response_schema"]=schema
        url=f"https://generativelanguage.googleapis.com/v1beta/models/{model}:generateContent?key={key}"
        try:
            req=urllib.request.Request(url,data=json.dumps(payload).encode(),headers={"Content-Type":"application/json"},method="POST")
            with urllib.request.urlopen(req,timeout=timeout) as r: data=json.loads(r.read().decode())
            text=data["candidates"][0]["content"]["parts"][0]["text"]
            return json.loads(text),idx,errors
        except urllib.error.HTTPError as e:
            errors.append(f"key {idx+1}: HTTP {e.code}")
            if e.code not in (400,401,403,408,409,429,500,502,503,504): break
        except Exception as e:
            errors.append(f"key {idx+1}: {e}")
    return None,-1,errors

def gemini_text_with_rotation(text: str, keys, model: str, instruction: str, timeout=25):
    """Text-only Gemini call with sequential key rotation on quota/rate-limit
    and transient failures. Returns (text, key_index, errors)."""
    if not text.strip() or not keys:
        return text, -1, ["No text or no API keys"]
    errors=[]
    for idx,key in enumerate(keys):
        payload={"contents":[{"parts":[{"text":instruction+"\n\nINPUT:\n"+text}]}],"generationConfig":{"temperature":0}}
        url=f"https://generativelanguage.googleapis.com/v1beta/models/{model}:generateContent?key={key}"
        try:
            req=urllib.request.Request(url,data=json.dumps(payload).encode(),headers={"Content-Type":"application/json"},method="POST")
            with urllib.request.urlopen(req,timeout=timeout) as r: data=json.loads(r.read().decode())
            result=data["candidates"][0]["content"]["parts"][0]["text"]
            return result.strip(),idx,errors
        except urllib.error.HTTPError as e:
            errors.append(f"key {idx+1}: HTTP {e.code}")
            if e.code not in (400,401,403,408,409,429,500,502,503,504): break
        except Exception as e:
            errors.append(f"key {idx+1}: {e}")
    return text,-1,errors
