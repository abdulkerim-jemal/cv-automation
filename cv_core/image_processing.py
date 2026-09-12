from __future__ import annotations

from dataclasses import dataclass
from typing import Literal
import math

import numpy as np
from PIL import Image, ImageDraw, ImageOps


@dataclass(frozen=True)
class CropBox:
    left: int
    top: int
    right: int
    bottom: int

    def as_tuple(self):
        return self.left, self.top, self.right, self.bottom

    @property
    def width(self):
        return max(0, self.right - self.left)

    @property
    def height(self):
        return max(0, self.bottom - self.top)


def clamp_box(box, width: int, height: int, min_size: int = 8):
    l, t, r, b = [int(x) for x in box]
    l, r = sorted((max(0, min(width, l)), max(0, min(width, r))))
    t, b = sorted((max(0, min(height, t)), max(0, min(height, b))))
    if r - l < min_size:
        r = min(width, l + min_size)
        l = max(0, r - min_size)
    if b - t < min_size:
        b = min(height, t + min_size)
        t = max(0, b - min_size)
    return l, t, r, b


def _edge_background(img: Image.Image):
    arr = np.asarray(img.convert("RGB"), dtype=np.int16)
    h, w = arr.shape[:2]
    strip = max(1, min(h, w) // 100)
    edge = np.concatenate([
        arr[:strip].reshape(-1, 3), arr[-strip:].reshape(-1, 3),
        arr[:, :strip].reshape(-1, 3), arr[:, -strip:].reshape(-1, 3),
    ])
    return tuple(int(x) for x in np.median(edge, axis=0))


def detect_content_bbox(img: Image.Image, tolerance: int = 18, min_content_frac: float = 0.05):
    rgb = img.convert("RGB")
    arr = np.asarray(rgb, dtype=np.int16)
    h, w = arr.shape[:2]
    if h < 4 or w < 4:
        return None
    bg = np.array(_edge_background(rgb), dtype=np.int16)
    diff = np.abs(arr - bg).sum(axis=2)
    mask = diff > tolerance * 3
    rows = np.where(mask.any(axis=1))[0]
    cols = np.where(mask.any(axis=0))[0]
    if not rows.size or not cols.size:
        return None
    top, bottom = rows[0], rows[-1]
    left, right = cols[0], cols[-1]
    pad = max(2, int(0.01 * max(h, w)))
    top, left = max(0, top-pad), max(0, left-pad)
    bottom, right = min(h-1, bottom+pad), min(w-1, right+pad)
    frac = ((bottom-top+1) * (right-left+1)) / float(w*h)
    if frac < min_content_frac:
        return None
    return left, top, right+1, bottom+1


def trim_whitespace(img: Image.Image, tolerance: int = 18, min_content_frac: float = 0.15):
    bbox = detect_content_bbox(img, tolerance=tolerance, min_content_frac=min_content_frac)
    return img.crop(bbox) if bbox else img.copy()


def _smart_vertical_crop(img: Image.Image, target_ratio: float, anchor: str):
    w, h = img.size
    current = w / h
    if current >= target_ratio:
        # Need to remove width; keep the whole vertical subject.
        new_w = max(1, int(round(h * target_ratio)))
        if anchor.endswith("left"):
            left = 0
        elif anchor.endswith("right"):
            left = w - new_w
        else:
            left = (w - new_w) // 2
        return img.crop((left, 0, left + new_w, h))
    # Need to remove height. Full-body images are top-biased by default so
    # head/upper body survive when the source is taller than the frame.
    new_h = max(1, int(round(w / target_ratio)))
    if anchor.startswith("top"):
        top = 0
    elif anchor.startswith("bottom"):
        top = h - new_h
    else:
        top = (h - new_h) // 2
    return img.crop((0, top, w, top + new_h))


def prepare_cv_image(source: Image.Image, target_width_px: int, target_height_px: int,
                     mode: Literal["cover", "contain"] = "contain",
                     anchor: str = "center", fill=(255, 255, 255)) -> Image.Image:
    """Return an exact-size RGB frame without distortion.

    contain keeps the whole photo, centered, with bars if the aspect differs.
    cover fills the frame by cropping overflow. CV photos use contain so
    passport / 3x4 / full-body images are never cropped inside their boxes.
    """
    target_width_px = max(1, int(target_width_px))
    target_height_px = max(1, int(target_height_px))
    src = source.convert("RGB")
    if mode == "cover":
        src = _smart_vertical_crop(src, target_width_px / target_height_px, anchor)
        return ImageOps.fit(src, (target_width_px, target_height_px), method=Image.Resampling.LANCZOS,
                            centering=(0.5, 0.0 if anchor.startswith("top") else 1.0 if anchor.startswith("bottom") else 0.5))
    fitted = src.copy()
    fitted.thumbnail((target_width_px, target_height_px), Image.Resampling.LANCZOS)
    canvas = Image.new("RGB", (target_width_px, target_height_px), fill)
    canvas.paste(fitted, ((target_width_px - fitted.width) // 2, (target_height_px - fitted.height) // 2))
    return canvas


def smart_crop_full_body(source: Image.Image, target_width_px: int, target_height_px: int):
    return prepare_cv_image(source, target_width_px, target_height_px, mode="contain", anchor="center")


def save_fixed_jpeg(source: Image.Image, path, width_px: int, height_px: int, anchor="top-center", quality=95):
    out = prepare_cv_image(source, width_px, height_px, mode="cover", anchor=anchor)
    out.save(path, "JPEG", quality=quality, optimize=True, progressive=True)
    return out


def default_boxes(w: int, h: int):
    return {
        "passport": (0, 0, int(w*0.50), h),
        "3x4": (int(w*0.51), int(h*0.03), int(w*0.67), int(h*0.32)),
        "full": (int(w*0.68), 0, w, h),
        "ocr": (0, int(h*0.60), int(w*0.50), h),
    }


def auto_detect_boxes(img: Image.Image):
    boxes = default_boxes(*img.size)
    out = {}
    for key, box in boxes.items():
        if key == "ocr":
            # Keep OCR inside the passport region. The composite also contains
            # a full-body photo, and including it creates OCR noise.
            pbox = out.get("passport", boxes["passport"])
            pl, pt, pr, pb = pbox
            out[key] = clamp_box(
                (pl, pt + int((pb - pt) * 0.55), pr, pb),
                *img.size,
            )
            continue
        crop = img.crop(box)
        content = detect_content_bbox(crop, tolerance=14, min_content_frac=0.05)
        if content is None:
            out[key] = box
        else:
            l,t,r,b = content
            out[key] = clamp_box((box[0]+l, box[1]+t, box[0]+r, box[1]+b), *img.size)
    return out


def draw_crop_preview(img: Image.Image, boxes, selected=None, width=900):
    scale = width / max(1, img.width)
    height = max(1, int(img.height*scale))
    bg = img.copy().resize((width, height), Image.Resampling.LANCZOS)
    d = ImageDraw.Draw(bg)
    for key, box in boxes.items():
        l,t,r,b = box
        color = "#ff8c00" if key == "ocr" else ("#00a86b" if key == selected else "#6b7280")
        d.rectangle((int(l*scale),int(t*scale),int(r*scale),int(b*scale)), outline=color, width=3)
    return bg
