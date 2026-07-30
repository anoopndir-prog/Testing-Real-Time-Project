"""Auto-crop helper for the seal identification photo (Final Test Report,
section 3.2). There is no true object-recognition model here - it is a
classical background-subtraction heuristic: it assumes the seal is
photographed against a reasonably plain, contrasting surface (a lab bench,
foam pad, sheet of paper, etc.) and finds the bounding box of whatever
stands out from that background. It is meant as a starting point the user
drags into place, not a guaranteed-exact crop - callers should always let
the user adjust the suggested box before it is baked into the report.
"""

from __future__ import annotations

import io

import numpy as np
from PIL import Image, ImageFilter, ImageOps

CropBox = tuple[int, int, int, int]  # left, top, right, bottom, in pixels

# Below this fraction of the frame, the "foreground" is treated as noise
# (nothing distinct found); above it, the background itself is treated as
# too busy/non-uniform for the heuristic to trust - both cases fall back
# to "no crop suggested" (the full photo) so the user draws the box by hand.
_MIN_FOREGROUND_COVERAGE = 0.005
_MAX_FOREGROUND_COVERAGE = 0.92
_ROW_COL_DENSITY_FLOOR = 0.04
_WORK_MAX_DIM = 900


def normalize_for_report(image: Image.Image) -> Image.Image:
    """Applies EXIF rotation (phone photos are frequently stored sideways
    with an orientation tag) and drops any alpha/CMYK channel."""
    return ImageOps.exif_transpose(image).convert("RGB")


def suggest_crop_box(image: Image.Image, padding_frac: float = 0.05) -> CropBox:
    """Returns a bounding box around whatever contrasts with the photo's
    background. Falls back to the full image when the background isn't
    uniform enough to separate confidently."""
    width, height = image.size
    if width < 20 or height < 20:
        return (0, 0, width, height)

    scale = min(1.0, _WORK_MAX_DIM / max(width, height))
    work = image.convert("L")
    if scale < 1.0:
        work = work.resize(
            (max(1, int(width * scale)), max(1, int(height * scale))), Image.LANCZOS
        )
    work = work.filter(ImageFilter.GaussianBlur(radius=2))
    arr = np.asarray(work, dtype=np.float32)
    work_h, work_w = arr.shape

    border = max(3, int(min(work_h, work_w) * 0.04))
    border_pixels = np.concatenate([
        arr[:border, :].ravel(),
        arr[-border:, :].ravel(),
        arr[:, :border].ravel(),
        arr[:, -border:].ravel(),
    ])
    background_value = float(np.median(border_pixels))
    background_noise = float(np.std(border_pixels))

    diff = np.abs(arr - background_value)
    diff_max = float(diff.max())
    if diff_max < 1e-3:
        return (0, 0, width, height)

    threshold = _otsu_threshold(diff.ravel(), diff_max)
    # A flat/textured background can still produce a low Otsu threshold;
    # require the foreground to clear the background's own noise level by
    # a comfortable margin, or everything (including background texture)
    # gets treated as "foreground".
    threshold = max(threshold, background_noise * 2.5, 8.0)

    mask = diff >= threshold
    coverage = float(mask.mean())
    if coverage < _MIN_FOREGROUND_COVERAGE or coverage > _MAX_FOREGROUND_COVERAGE:
        return (0, 0, width, height)

    left, right = _density_bounds(mask.mean(axis=0))
    top, bottom = _density_bounds(mask.mean(axis=1))

    box_w = right - left
    box_h = bottom - top
    pad_x = int(box_w * padding_frac) + 2
    pad_y = int(box_h * padding_frac) + 2
    left = max(0, left - pad_x)
    top = max(0, top - pad_y)
    right = min(work_w, right + pad_x)
    bottom = min(work_h, bottom + pad_y)

    if scale < 1.0:
        inv = 1.0 / scale
        left, top, right, bottom = (
            int(left * inv), int(top * inv), int(right * inv), int(bottom * inv)
        )

    left = max(0, min(left, width - 1))
    top = max(0, min(top, height - 1))
    right = max(left + 1, min(right, width))
    bottom = max(top + 1, min(bottom, height))
    return (left, top, right, bottom)


def _density_bounds(density: np.ndarray) -> tuple[int, int]:
    above = np.where(density > _ROW_COL_DENSITY_FLOOR)[0]
    if above.size == 0:
        return 0, density.size
    return int(above[0]), int(above[-1]) + 1


def _otsu_threshold(values: np.ndarray, value_max: float) -> float:
    """Otsu's method over a 256-bin histogram of `values` (range 0..value_max),
    returned back in the original value units."""
    if value_max <= 0:
        return 0.0
    hist, _ = np.histogram(values, bins=256, range=(0.0, value_max))
    hist = hist.astype(np.float64)
    total = hist.sum()
    if total == 0:
        return value_max / 2

    bin_centers = np.arange(256)
    sum_all = float(np.dot(bin_centers, hist))
    weight_bg = 0.0
    sum_bg = 0.0
    best_variance = 0.0
    best_bin = 128
    for bin_idx in range(256):
        weight_bg += hist[bin_idx]
        if weight_bg == 0:
            continue
        weight_fg = total - weight_bg
        if weight_fg == 0:
            break
        sum_bg += bin_idx * hist[bin_idx]
        mean_bg = sum_bg / weight_bg
        mean_fg = (sum_all - sum_bg) / weight_fg
        variance_between = weight_bg * weight_fg * (mean_bg - mean_fg) ** 2
        if variance_between > best_variance:
            best_variance = variance_between
            best_bin = bin_idx
    return (best_bin / 255.0) * value_max


def crop_to_box(image: Image.Image, box: CropBox) -> Image.Image:
    left, top, right, bottom = box
    width, height = image.size
    left = max(0, min(int(left), width - 1))
    top = max(0, min(int(top), height - 1))
    right = max(left + 1, min(int(right), width))
    bottom = max(top + 1, min(int(bottom), height))
    return image.crop((left, top, right, bottom))


def encode_jpeg(image: Image.Image, quality: int = 90) -> bytes:
    buffer = io.BytesIO()
    image.convert("RGB").save(buffer, format="JPEG", quality=quality, optimize=True)
    return buffer.getvalue()
