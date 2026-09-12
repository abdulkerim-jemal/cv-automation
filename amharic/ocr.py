from __future__ import annotations

import csv
import os
import re
import subprocess
import tempfile
from pathlib import Path

import numpy as np
from PIL import Image, ImageOps, ImageFilter


# Ge'ez syllabary + punctuation commonly encountered in Ethiopian documents.
GE_EZ_CHARS = (
    "፩፪፫፬፭፮፯፰፱"
    "ሀሁሂሃሄህሆለሉሊላሌልሎሐሑሒሓሔሕሖ"
    "መሙሚማሜምሞሠሡሢሣሤሥሦረሩሪራሬርሮሰሱሲሳሴስሶሸሹሺሻሼሽሾ"
    "ቀቁቂቃቄቅቆበቡቢባቤብቦቨቩቪቫቬቭቮተቱቲታቴትቶቸቹቺቻቼችቾ"
    "ኀኁኂኃኄኅኆነኑኒናኔንኖኘኙኚኛኜኝኞአኡኢኣኤእኦከኩኪካኬክኮ"
    "ኸኹኺኻኼኽኾወዉዊዋዌውዎዐዑዒዓዔዕዖዘዙዚዛዜዝዞዠዡዢዣዤዥዦ"
    "የዩዪያዬይዮደዱዲዳዴድዶጀጁጂጃጄጅጆገጉጊጋጌግጎጠጡጢጣጤጥጦጨጩጪጫጬጭጮ"
    "ጰጱጲጳጴጵጶጸጹጺጻጼጽጾፀፁፂፃፄፅፆፈፉፊፋፌፍፎፐፑፒፓፔፕፖ"
    "፣።፤፥፦፧፨?!,:;.()-/\"' 0123456789"
)
VOCAB = {c:i+1 for i,c in enumerate(dict.fromkeys(GE_EZ_CHARS))}
VOCAB["<BLANK>"]=0
INV_VOCAB={v:k for k,v in VOCAB.items()}


def tesseract_amharic(image: Image.Image, psm=6):
    try:
        import pytesseract
    except Exception as e:
        return "", f"pytesseract unavailable: {e}"
    try:
        from ocr.pipeline import find_tesseract
        cmd,_=find_tesseract()
        if not cmd: return "", "Tesseract executable unavailable"
        pytesseract.pytesseract.tesseract_cmd=cmd
        gray=ImageOps.autocontrast(image.convert("L"),cutoff=1).filter(ImageFilter.SHARPEN)
        lang="amh"
        try:
            langs=pytesseract.get_languages(config="")
            if "amh" not in langs and "Ethiopic" in langs: lang="Ethiopic"
        except Exception: pass
        text=pytesseract.image_to_string(gray,lang=lang,config=f"--oem 3 --psm {psm}")
        return text, f"Tesseract 5.x ({lang})"
    except Exception as e:
        return "", f"Tesseract Amharic OCR failed: {e}"


def preprocess_amharic(image: Image.Image):
    gray=image.convert("L")
    if gray.width<1600:
        scale=1600/max(1,gray.width)
        gray=gray.resize((int(gray.width*scale),int(gray.height*scale)),Image.Resampling.LANCZOS)
    gray=ImageOps.autocontrast(gray,cutoff=1)
    return gray.filter(ImageFilter.SHARPEN)


def segment_lines(image: Image.Image):
    """Simple horizontal projection segmentation for optional CRNN inference."""
    arr=np.asarray(preprocess_amharic(image))
    ink=255-arr
    projection=(ink>35).sum(axis=1)
    active=projection>max(2,int(arr.shape[1]*0.002))
    spans=[]; start=None
    for i,v in enumerate(active):
        if v and start is None: start=i
        if not v and start is not None:
            if i-start>=4: spans.append((start,i))
            start=None
    if start is not None: spans.append((start,len(active)))
    return [image.crop((0,max(0,t-4),image.width,min(image.height,b+4))) for t,b in spans]


def postprocess_with_gemini(text: str, keys, model="gemini-2.5-flash"):
    if not text.strip() or not keys:
        return text,{"used":False,"reason":"No text or no Gemini keys"}
    from ocr.pipeline import gemini_text_with_rotation
    instruction=("Correct OCR errors in the following Amharic/Ge'ez text. Preserve names, numbers, "
                 "line breaks where possible, and meaning. Do not invent missing content. Return only the corrected text.")
    corrected,idx,errors=gemini_text_with_rotation(text,keys,model,instruction)
    return corrected,{"used":idx>=0,"key_index":idx,"errors":errors}


# ---- Custom CNN + BiLSTM + CTC (not a pretrained HF model) -----------------
def build_crnn(num_classes: int|None=None, img_height: int=48, hidden: int=128):
    import torch
    import torch.nn as nn
    classes=num_classes or len(VOCAB)
    class CRNN(nn.Module):
        def __init__(self):
            super().__init__()
            self.cnn=nn.Sequential(
                nn.Conv2d(1,64,3,padding=1),nn.BatchNorm2d(64),nn.ReLU(True),nn.MaxPool2d(2,2),
                nn.Conv2d(64,128,3,padding=1),nn.BatchNorm2d(128),nn.ReLU(True),nn.MaxPool2d(2,2),
                nn.Conv2d(128,256,3,padding=1),nn.BatchNorm2d(256),nn.ReLU(True),
                nn.Conv2d(256,256,3,padding=1),nn.BatchNorm2d(256),nn.ReLU(True),nn.MaxPool2d((2,1),(2,1)),
                nn.Conv2d(256,512,3,padding=1),nn.BatchNorm2d(512),nn.ReLU(True),nn.MaxPool2d((2,1),(2,1)),
            )
            self.rnn=nn.LSTM(512,hidden,num_layers=2,bidirectional=True,dropout=0.1)
            self.fc=nn.Linear(hidden*2,classes)
        def forward(self,x):
            x=self.cnn(x)
            x=x.mean(dim=2).permute(2,0,1)
            x,_=self.rnn(x)
            return self.fc(x)
    return CRNN()


def greedy_ctc_decode(logits):
    import torch
    ids=logits.argmax(-1).detach().cpu().numpy()
    out=[]
    for seq in ids.T if ids.ndim==2 else ids:
        prev=0; chars=[]
        for i in seq:
            if i!=0 and i!=prev: chars.append(INV_VOCAB.get(int(i),""))
            prev=i
        out.append("".join(chars))
    return out


def load_crnn_checkpoint(path, device="cpu"):
    import torch
    model=build_crnn()
    state=torch.load(path,map_location=device)
    model.load_state_dict(state.get("model",state))
    model.to(device).eval()
    return model


def crnn_infer_lines(model, lines, device="cpu", width=1024, height=48):
    import torch
    batch=[]
    for im in lines:
        g=preprocess_amharic(im).resize((width,height),Image.Resampling.LANCZOS)
        arr=np.asarray(g,dtype=np.float32)/255.0
        batch.append(arr[None,None])
    if not batch:return []
    x=torch.from_numpy(np.concatenate(batch,axis=0)).to(device)
    with torch.no_grad(): logits=model(x)
    return greedy_ctc_decode(logits)


def train_from_csv(csv_path, image_root, output_path, epochs=20, batch_size=8, lr=1e-3, device="cpu"):
    """Train from CSV columns: image,text. This is a real CTC training path,
    intentionally independent of Hugging Face pretrained models."""
    import torch
    from torch.utils.data import Dataset,DataLoader
    import torch.nn as nn
    class DS(Dataset):
        def __init__(self):
            self.rows=[]
            with open(csv_path,encoding="utf-8") as f:
                for r in csv.DictReader(f): self.rows.append(r)
        def __len__(self): return len(self.rows)
        def __getitem__(self,i):
            r=self.rows[i]; im=Image.open(Path(image_root)/r["image"]).convert("L")
            im=preprocess_amharic(im).resize((512,48),Image.Resampling.LANCZOS)
            x=torch.from_numpy(np.asarray(im,dtype=np.float32)/255.).unsqueeze(0)
            y=torch.tensor([VOCAB.get(c,0) for c in r["text"] if c in VOCAB],dtype=torch.long)
            return x,y
    ds=DS()
    def collate(batch):
        xs=[b[0] for b in batch]; ys=[b[1] for b in batch]
        return torch.stack(xs),torch.cat(ys),torch.tensor([len(y) for y in ys],dtype=torch.long)
    loader=DataLoader(ds,batch_size=batch_size,shuffle=True,collate_fn=collate)
    model=build_crnn().to(device); opt=torch.optim.AdamW(model.parameters(),lr=lr); ctc=nn.CTCLoss(blank=0,zero_infinity=True)
    for epoch in range(epochs):
        model.train(); total=0.0
        for x,y,lens in loader:
            x,y=x.to(device),y.to(device); logits=model(x); T=logits.size(0)
            inp_lens=torch.full((x.size(0),),T,dtype=torch.long,device=device)
            loss=ctc(logits.log_softmax(2),y,inp_lens,lens.to(device))
            opt.zero_grad(); loss.backward(); torch.nn.utils.clip_grad_norm_(model.parameters(),5); opt.step(); total+=float(loss)
        print(f"epoch {epoch+1}/{epochs}: loss={total/max(1,len(loader)):.4f}")
    torch.save({"model":model.state_dict(),"vocab":VOCAB},output_path)
    return output_path
