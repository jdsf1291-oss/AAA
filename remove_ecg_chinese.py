#!/usr/bin/env python3
"""Remove Chinese labels from an ECG printout, replacing them with English."""

from pathlib import Path

import numpy as np
from PIL import Image, ImageDraw, ImageFont

SRC = Path(__file__).resolve().parent / "ecg_original.png"
OUT = Path(__file__).resolve().parent / "ecg_no_chinese.png"
FONT_PATH = "/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf"


def clone_horizontal(img: np.ndarray, box, src_x0: int, period: int) -> None:
    """Cover dest box with same-row pixels starting at src_x0, tiled by period.

    Using the same Y as the destination keeps ECG-paper grid phase (including
    screenshot drift). period must be a multiple of the 30px major square.
    """
    x0, y0, x1, y1 = [int(v) for v in box]
    h, w = img.shape[:2]
    x0, y0 = max(0, x0), max(0, y0)
    x1, y1 = min(w, x1), min(h, y1)
    band = img[y0:y1, :, :].copy()
    for x in range(x0, x1):
        sx = src_x0 + ((x - x0) % period)
        if 0 <= sx < w:
            img[y0:y1, x] = band[:, sx]


def main() -> None:
    img = np.array(Image.open(SRC).convert("RGB"))

    # Lines 1-5 live on the white header. Stop above the 床号 / P-R-T row.
    img[0:186, :] = (255, 255, 255)

    # Same-row clone-stamp so the pink grid stays aligned.
    # Gap between the left patient column and the middle measurements (~x=280-450).
    clone_horizontal(img, (62, 188, 158, 216), src_x0=302, period=120)   # 床号:003
    clone_horizontal(img, (455, 188, 775, 216), src_x0=305, period=120)  # P-R-T电轴 line
    clone_horizontal(img, (820, 196, 972, 226), src_x0=308, period=120)  # 年龄不确定
    clone_horizontal(img, (828, 226, 956, 258), src_x0=308, period=120)  # 异常 ECG
    # Nearby empty paper, 4 major squares left of the 梁玲玉 watermark.
    # Direct copy (no tiling) so the last column is not wrapped.
    img[956:980, 1488:1590] = img[956:980, 1368:1470]
    img[980:991, 1488:1590] = (255, 255, 255)

    im = Image.fromarray(img)
    draw = ImageDraw.Draw(im)
    font = ImageFont.truetype(FONT_PATH, 16)
    font_sm = ImageFont.truetype(FONT_PATH, 15)
    ink = (30, 30, 30)

    left_x, mid_x, right_x = 68, 459, 793
    y1, y2, y3, y4, y5 = 16, 51, 86, 122, 157

    draw.text((left_x, y1), "ID: 0021457665", font=font, fill=ink)
    draw.text((left_x, y2), "Name: Wu Qingyu", font=font, fill=ink)
    draw.text((left_x, y3), "Sex: M", font=font, fill=ink)
    draw.text((left_x, y4), "Age: 16", font=font, fill=ink)
    draw.text((left_x, y5), "Dept: Cardiology", font=font, fill=ink)

    draw.text((mid_x, y1), "HR(60-105): 57 bpm  \u2193", font=font, fill=ink)
    draw.text((mid_x, y2), "PR interval: - ms", font=font, fill=ink)
    draw.text((mid_x, y3), "QRS(~110): 180 ms  \u2191", font=font, fill=ink)
    draw.text((mid_x, y4), "QT/QTc(360-450): 502/489 ms  \u2191", font=font, fill=ink)
    draw.text((mid_x, y5), "RV5/SV1(~4): 0.781/0.336mV(R+S:1.117mV)", font=font_sm, fill=ink)

    draw.text((right_x, y1), "Diagnosis:", font=font, fill=ink)

    draw.text((left_x, 194), "Bed: 003", font=font, fill=ink)
    draw.text((mid_x, 193), "P-R-T axis(-30-90): -/180/-23\u00b0  \u2191", font=font, fill=ink)
    draw.text((848, 204), "Age uncertain", font=font_sm, fill=ink)
    draw.text((828, 232), "Abnormal ECG", font=font_sm, fill=ink)

    im.save(OUT, "PNG")
    print(f"Wrote {OUT} {im.size}")


if __name__ == "__main__":
    main()
