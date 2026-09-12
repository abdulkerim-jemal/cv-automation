from __future__ import annotations

import base64
from io import BytesIO
from pathlib import Path

from PIL import Image
import streamlit as st
import streamlit.components.v1 as components

_FRONTEND = Path(__file__).resolve().parent / "crop_canvas_frontend"
_crop_canvas = components.declare_component("cv_drag_crop", path=str(_FRONTEND))


def image_to_data_url(image: Image.Image, width: int, height: int) -> str:
    preview = image.convert("RGB").resize((max(1, width), max(1, height)), Image.Resampling.LANCZOS)
    buf = BytesIO()
    preview.save(buf, format="JPEG", quality=82, optimize=True)
    encoded = base64.b64encode(buf.getvalue()).decode("ascii")
    return f"data:image/jpeg;base64,{encoded}"


def drag_crop_box(
    image: Image.Image,
    current_box: tuple[int, int, int, int],
    color: str,
    key: str,
    max_width: int = 720,
    max_height: int = 2000,  # increased from 480 to show full image
) -> tuple[int, int, int, int]:
    """Interactive drag rectangle on the source image. Returns original-pixel box."""
    orig_w, orig_h = image.size
    if orig_w < 8 or orig_h < 8:
        return current_box

    scale = min(max_width / orig_w, max_height / orig_h, 1.0)
    disp_w = max(1, int(round(orig_w * scale)))
    disp_h = max(1, int(round(orig_h * scale)))
    data_url = image_to_data_url(image, disp_w, disp_h)

    left, top, right, bottom = current_box
    display_box = [
        left * scale,
        top * scale,
        right * scale,
        bottom * scale,
    ]
    result = _crop_canvas(
        image_data_url=data_url,
        width=disp_w,
        height=disp_h,
        color=color,
        box=display_box,
        key=key,
        default=None,
    )
    if not result or not isinstance(result, dict) or "left" not in result:
        return current_box

    ts = result.get("_ts", 0)
    ts_key = f"_canvas_handled_ts_{key}"
    last_ts = st.session_state.get(ts_key, 0)
    if ts and ts == last_ts:
        return current_box
    if ts:
        st.session_state[ts_key] = ts

    new_left = int(round(result["left"] / scale))
    new_top = int(round(result["top"] / scale))
    new_right = int(round(result["right"] / scale))
    new_bottom = int(round(result["bottom"] / scale))
    new_left = max(0, min(orig_w, new_left))
    new_top = max(0, min(orig_h, new_top))
    new_right = max(0, min(orig_w, new_right))
    new_bottom = max(0, min(orig_h, new_bottom))
    if new_right - new_left < 8 or new_bottom - new_top < 8:
        return current_box
    return (new_left, new_top, new_right, new_bottom)