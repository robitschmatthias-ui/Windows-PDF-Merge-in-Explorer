"""PDF operations: convert images, merge, rotate, fit to A4.

Mirrors the pdftk / img2pdf / Ghostscript steps of the Linux version
(Linux-PDF-Merge-in-Nemo), implemented with pypdf and img2pdf.
"""

from __future__ import annotations

import io
import os
from pathlib import Path

import img2pdf
from PIL import Image, ImageSequence
from pypdf import PdfReader, PdfWriter, Transformation
from pypdf.generic import RectangleObject

A4_W = 595.28
A4_H = 841.89
A4_LONG_IN = 297 / 25.4
A4_SHORT_IN = 210 / 25.4

# Advanced "fit to A4": images reaching at least this resolution on a full A4
# page fill the page; lower-resolution images keep their original size.
MIN_FILL_DPI = 200

PDF_EXT = {".pdf"}
IMAGE_EXT = {".jpg", ".jpeg", ".png", ".gif", ".tif", ".tiff", ".bmp", ".webp"}


class PdfError(Exception):
    """Error with a message meant for the user."""


# ── Images ──────────────────────────────────────────────────────────────────

def _flatten_frames(src: Path) -> list[bytes]:
    """Renders every frame onto a white background (RGB) and returns PNGs.

    Fallback for images img2pdf refuses directly (e.g. alpha channel).
    """
    frames = []
    with Image.open(src) as img:
        for frame in ImageSequence.Iterator(img):
            frame = frame.convert("RGBA")
            background = Image.new("RGB", frame.size, "white")
            background.paste(frame, mask=frame.getchannel("A"))
            buf = io.BytesIO()
            background.save(buf, format="PNG")
            frames.append(buf.getvalue())
    return frames


def _a4_fit_mode(src: Path, mode: str):
    if mode == "fill":
        return img2pdf.FitMode.into
    # "keep-small": fill A4 only if the image still has MIN_FILL_DPI there,
    # otherwise keep its original size (shrink only if larger than A4)
    try:
        with Image.open(src) as img:
            w, h = img.size
        dpi = min(max(w, h) / A4_LONG_IN, min(w, h) / A4_SHORT_IN)
    except Exception:
        return img2pdf.FitMode.into
    return img2pdf.FitMode.into if dpi >= MIN_FILL_DPI else img2pdf.FitMode.shrink


def image_to_pdf(src: Path, dst: Path, a4: str | None = None) -> None:
    """Converts an image to PDF.

    a4=None:          page has the image's size (like img2pdf on Linux)
    a4="fill":        fitted to A4, portrait/landscape follows the image
    a4="keep-small":  A4 page; the image fills it only at >= MIN_FILL_DPI,
                      otherwise it stays at original size, centered
    """
    kwargs = {}
    if a4:
        kwargs["layout_fun"] = img2pdf.get_layout_fun(
            (A4_W, A4_H), fit=_a4_fit_mode(src, a4), auto_orient=True
        )
    try:
        data = img2pdf.convert(str(src), **kwargs)
    except Exception:
        try:
            data = img2pdf.convert(*_flatten_frames(src), **kwargs)
        except Exception as exc:
            raise PdfError(f"cannot read image ({exc})") from exc
    dst.write_bytes(data)


def is_image(path: Path) -> bool:
    """Detects images even without a matching extension (like 'file' on Linux)."""
    try:
        with Image.open(path) as img:
            img.verify()
        return True
    except Exception:
        return False


def is_pdf(path: Path) -> bool:
    try:
        with open(path, "rb") as f:
            return f.read(1024).lstrip().startswith(b"%PDF")
    except OSError:
        return False


# ── Merge ───────────────────────────────────────────────────────────────────

def open_pdf(path: Path) -> PdfReader:
    """Reads the whole PDF into memory.

    Holds no file lock, so the result may replace one of the source files
    (Advanced suggests the name of the first file).
    """
    try:
        reader = PdfReader(io.BytesIO(path.read_bytes()))
    except Exception as exc:
        raise PdfError(f"cannot read PDF ({exc})") from exc
    if reader.is_encrypted:
        # PDFs with permissions-only protection (no open password) decrypt
        # with an empty password; PDFs with an open password do not.
        try:
            ok = reader.decrypt("")
        except Exception:
            ok = 0
        if not ok:
            raise PdfError("password protected")
    return reader


def merge(readers: list[PdfReader]) -> PdfWriter:
    writer = PdfWriter()
    for reader in readers:
        writer.append(reader)
    return writer


def write_atomic(writer: PdfWriter, dst: Path) -> None:
    """Writes to a sibling file first, then replaces the target.

    If writing fails, an existing target stays untouched and no half-written
    PDF is left behind.
    """
    tmp = dst.with_name(f".{dst.name}.tmp")
    try:
        with open(tmp, "wb") as f:
            writer.write(f)
        os.replace(tmp, dst)
    finally:
        if tmp.exists():
            tmp.unlink()


# ── Orientation and A4 ──────────────────────────────────────────────────────

def _effective_size(page) -> tuple[float, float]:
    w = float(page.mediabox.width)
    h = float(page.mediabox.height)
    if page.rotation % 180 == 90:
        w, h = h, w
    return w, h


def _fit_to(page, target_w: float, target_h: float) -> None:
    """Scales the page to target_w x target_h and centers it
    (like Ghostscript -dPDFFitPage -dFIXEDMEDIA on Linux)."""
    page.transfer_rotation_to_content()
    box = page.mediabox
    w, h = float(box.width), float(box.height)
    scale = min(target_w / w, target_h / h)
    tx = (target_w - w * scale) / 2 - float(box.left) * scale
    ty = (target_h - h * scale) / 2 - float(box.bottom) * scale
    page.add_transformation(Transformation().scale(scale, scale).translate(tx, ty))
    rect = RectangleObject([0, 0, target_w, target_h])
    page.mediabox = rect
    page.cropbox = rect
    for key in ("/TrimBox", "/BleedBox", "/ArtBox"):
        if key in page:
            del page[key]


def postprocess(writer: PdfWriter, size: str, orient: str, report=None) -> None:
    """Adjusts page size and orientation.

    size:   "original" | "a4"
    orient: "original" | "portrait" | "landscape"

    With a fixed orientation, pages in the wrong orientation are rotated by
    90° (relative); with A4 they are then fitted into A4 in that orientation.
    Pages that already are A4 (images placed by image_to_pdf) keep scale 1.
    """
    if size == "original" and orient == "original":
        return
    total = len(writer.pages)
    for i, page in enumerate(writer.pages, start=1):
        if orient != "original":
            w, h = _effective_size(page)
            if (orient == "landscape") != (w > h):
                page.rotate(90)
        if size == "a4":
            if orient == "original":
                w, h = _effective_size(page)
                landscape = w > h
            else:
                landscape = orient == "landscape"
            if landscape:
                _fit_to(page, A4_H, A4_W)
            else:
                _fit_to(page, A4_W, A4_H)
        if report:
            report(i, total)
