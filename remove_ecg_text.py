#!/usr/bin/env python3
"""Erase all text from an ECG printout, keeping only waveforms on the red grid."""

from pathlib import Path

import numpy as np
from PIL import Image

SRC = Path(__file__).resolve().parent / "ecg_original.png"
OUT = Path(__file__).resolve().parent / "ecg_waveforms_only.png"

GRID_TOP = 203

# Original-image coordinates.
TEXT_BOXES = [
    (50, 184, 175, 222),       # 床号:003
    (430, 184, 800, 222),      # P-R-T电轴
    (805, 192, 990, 234),      # 年龄不确定
    (795, 216, 980, 268),      # 异常 ECG
    (56, 296, 105, 345),       # I
    (52, 405, 115, 460),       # II
    (52, 518, 122, 580),       # III
    (55, 642, 115, 700),       # V1
    (52, 758, 115, 820),       # II rhythm
    (48, 878, 115, 942),       # V5 rhythm
    (415, 296, 515, 355),      # aVR
    (415, 412, 515, 475),      # aVL
    (415, 528, 520, 590),      # aVF
    (785, 390, 870, 452),      # V2
    (785, 515, 872, 582),      # V3
    (1165, 284, 1255, 346),    # V4
    (1165, 404, 1255, 466),    # V5
    (1165, 524, 1255, 586),    # V6
]

WATERMARK = (1480, 948, 1590, 991)


def achromatic_ink(rgb: np.ndarray) -> np.ndarray:
    r = rgb[:, :, 0].astype(np.int16)
    g = rgb[:, :, 1].astype(np.int16)
    b = rgb[:, :, 2].astype(np.int16)
    gray = (r + g + b) / 3.0
    return (np.abs(r - g) < 42) & (np.abs(r - b) < 42) & (gray < 205)


def is_paper(rgb: np.ndarray) -> np.ndarray:
    """White paper or pink/red grid lines — not ink and not gray text fringes."""
    r = rgb[:, :, 0].astype(np.int16)
    g = rgb[:, :, 1].astype(np.int16)
    b = rgb[:, :, 2].astype(np.int16)
    white = (r > 248) & (g > 248) & (b > 248)
    pink = (r > 200) & (g > 150) & ((r - g) > 6)
    return white | pink


def column_span(ink_col: np.ndarray) -> int:
    ys = np.flatnonzero(ink_col)
    if ys.size == 0:
        return 0
    return int(ys.max() - ys.min() + 1)


def cover_with_grid(img: np.ndarray, orig: np.ndarray, paper: np.ndarray, box) -> None:
    """Rebuild pink grid in the box from nearby paper pixels on the same rows."""
    h, w = paper.shape
    x0, y0, x1, y1 = [int(v) for v in box]
    x0, y0 = max(0, x0), max(0, y0)
    x1, y1 = min(w, x1), min(h, y1)
    if x1 <= x0 or y1 <= y0:
        return
    period = 30
    xa, xb = max(0, x0 - 420), min(w, x1 + 420)
    for y in range(y0, y1):
        tmpl = np.zeros((period, 3), dtype=np.float64)
        cnt = np.zeros(period, dtype=np.float64)
        for x in range(xa, xb):
            if x0 <= x < x1:
                continue
            if paper[y, x]:
                k = x % period
                tmpl[k] += orig[y, x]
                cnt[k] += 1
        if (cnt == 0).any():
            for x in range(w):
                if paper[y, x]:
                    k = x % period
                    if cnt[k] == 0:
                        tmpl[k] += orig[y, x]
                        cnt[k] += 1
        cnt[cnt == 0] = 1
        tmpl = (tmpl / cnt[:, None]).astype(np.uint8)
        xs = np.arange(x0, x1)
        img[y, x0:x1] = tmpl[xs % period]


def thin_ys(ink_col: np.ndarray, gray_col: np.ndarray) -> np.ndarray | None:
    ys = np.flatnonzero(ink_col)
    if ys.size < 2 or (ys.max() - ys.min() + 1) > 9:
        return None
    if ys.size <= 3:
        return ys
    dark = 255 - gray_col[ys]
    pick = ys[np.argsort(dark)[-3:]]
    return np.sort(pick)


def restore_trace(
    img: np.ndarray,
    orig: np.ndarray,
    ink: np.ndarray,
    paper: np.ndarray,
    gray: np.ndarray,
    box,
    pad: int = 12,
) -> None:
    h, w = ink.shape
    x0, y0, x1, y1 = [int(v) for v in box]
    x0, y0 = max(0, x0), max(0, y0)
    x1, y1 = min(w, x1), min(h, y1)
    if x1 - x0 < 2 or y1 - y0 < 2:
        return

    cover_with_grid(img, orig, paper, (x0, y0, x1, y1))

    ref_x, ref_ys = None, None
    for x in range(x1, min(w, x1 + pad)):
        ys = thin_ys(ink[y0:y1, x], gray[y0:y1, x])
        if ys is not None:
            ref_x, ref_ys = x, ys
            break
    if ref_ys is None:
        for x in range(x0 - 1, max(-1, x0 - pad - 1), -1):
            ys = thin_ys(ink[y0:y1, x], gray[y0:y1, x])
            if ys is not None:
                ref_x, ref_ys = x, ys
                break
    if ref_ys is not None:
        colors = orig[y0:y1, ref_x]
        for x in range(x0, x1):
            img[y0 + ref_ys, x] = colors[ref_ys]

    for x in list(range(x0, min(x1, x0 + 8))) + list(range(max(x0, x1 - 8), x1)):
        if column_span(ink[y0:y1, x]) >= 12:
            img[y0:y1, x] = orig[y0:y1, x]


def main() -> None:
    orig_full = np.array(Image.open(SRC).convert("RGB"))
    ink_full = achromatic_ink(orig_full)
    paper_full = is_paper(orig_full)

    orig = orig_full[GRID_TOP:].copy()
    img = orig.copy()
    ink = ink_full[GRID_TOP:].copy()
    paper = paper_full[GRID_TOP:].copy()
    gray = orig.mean(axis=2)
    dy = GRID_TOP

    # First grid rows only held header text; rebuild them as paper.
    cover_with_grid(img, orig, paper, (0, 0, img.shape[1], 18))

    for x0, y0, x1, y1 in TEXT_BOXES:
        box = (x0, y0 - dy, x1, y1 - dy)
        if box[3] <= 0 or box[1] >= img.shape[0]:
            continue
        restore_trace(img, orig, ink, paper, gray, box)

    x0, y0, x1, y1 = WATERMARK
    y0, y1 = max(0, y0 - dy), min(img.shape[0], y1 - dy)
    src_x0 = x0 - 120
    img[y0:y1, x0:x1] = orig[y0:y1, src_x0 : src_x0 + (x1 - x0)]

    Image.fromarray(img).save(OUT, "PNG")
    print(f"Wrote {OUT} {img.shape[1]}x{img.shape[0]}")


if __name__ == "__main__":
    main()
