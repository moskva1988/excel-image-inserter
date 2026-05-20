"""Image Batch Processor — folder-to-folder image transformation worker.

Operates entirely independently of the Excel pipeline. Reads a folder of
images, optionally resizes / crops / watermarks them, writes them out to a
destination folder (or overwrites in place), then patches the file metadata
(atime / mtime, optionally EXIF DateTimeOriginal, optionally Windows creation
time).

The worker is a `QThread`; UI code listens to `progress` / `finished` signals
exactly the way `InsertWorker` is used by the Excel tab.
"""

from __future__ import annotations

import os
import platform
from datetime import datetime
from pathlib import Path

from PyQt5.QtCore import QThread, pyqtSignal

from PIL import Image as PILImage, ImageDraw, ImageFont


# ── Date formats (label, strftime pattern, uppercase_month_flag) ──────────────
# The uppercase flag tells the renderer to .upper() the strftime output so
# patterns like "%d/%b/%Y" yield "13/JAN/2025" rather than "13/Jan/2025".
DATE_FORMATS = [
    ("2024/10/06 (YYYY/MM/DD)",            "%Y/%m/%d",         False),
    ("06/10/2024 (DD/MM/YYYY — EU)",       "%d/%m/%Y",         False),
    ("10/06/2024 (MM/DD/YYYY — US)",       "%m/%d/%Y",         False),
    ("2024-10-06 (ISO)",                   "%Y-%m-%d",         False),
    ("06.10.2024 (DD.MM.YYYY)",            "%d.%m.%Y",         False),
    ("13/JAN/2025 (DD/MMM/YYYY uppercase)", "%d/%b/%Y",        True),
    ("13-Jan-2025 (DD-MMM-YYYY)",          "%d-%b-%Y",         False),
    ("Jan 13, 2025 (US short)",            "%b %d, %Y",        False),
    ("September 13, 2024 (Full month)",    "%B %d, %Y",        False),
    ("13 September 2024",                  "%d %B %Y",         False),
    ("September 13, 2024 10:51 AM",        "%B %d, %Y %I:%M %p", False),
    ("2024-10-06 10:51 (ISO + 24h)",       "%Y-%m-%d %H:%M",   False),
    ("06.10.2024 10:51 (EU + 24h)",        "%d.%m.%Y %H:%M",   False),
]


SUPPORTED_EXT = {".jpg", ".jpeg", ".png", ".bmp", ".tif", ".tiff", ".webp"}


# Watermark position keys exposed to the UI dropdown.
POSITIONS = [
    "Bottom-right", "Bottom-left", "Bottom-center",
    "Top-right", "Top-left", "Top-center",
    "Center",
]


COLOR_PRESETS = {
    "White":      (255, 255, 255),
    "Black":      (0,   0,   0),
    "Orange":     (255, 140, 0),
    "Yellow":     (255, 220, 0),
    "Red":        (220, 30,  30),
    "Cyan":       (0,   200, 220),
    "Lime green": (120, 220, 80),
}


def format_date(dt: datetime, fmt_index: int) -> str:
    """Render `dt` using the format at `fmt_index` from DATE_FORMATS."""
    if fmt_index < 0 or fmt_index >= len(DATE_FORMATS):
        fmt_index = 0
    _label, pattern, upper_month = DATE_FORMATS[fmt_index]
    out = dt.strftime(pattern)
    if upper_month:
        out = out.upper()
    return out


# ── Worker thread ─────────────────────────────────────────────────────────────
class BatchProcessorWorker(QThread):
    """Processes a folder of images in a background thread.

    Config dict shape (all keys optional unless noted):
      input_dir            (str, required)
      output_dir           (str)  — required unless overwrite=True
      overwrite            (bool) — write back to source file
      resize_mode          ("none" | "long_side" | "percent" | "exact")
      resize_long_side     (int)
      resize_percent       (int)
      resize_w, resize_h   (int)
      resize_keep_aspect   (bool)
      crop_ratio           ((w, h) tuple or None)
      created_dt           (datetime or None)
      modified_dt          (datetime or None)
      write_exif           (bool)
      watermark_mode       ("none" | "date" | "image")
      wm_date_dt           (datetime)
      wm_date_format_index (int)
      wm_color             ((r, g, b) tuple)
      wm_font              (str — family name; falls back to Pillow default)
      wm_font_size_pct     (int 1..20)  — % of image width
      wm_position          (str — one of POSITIONS)
      wm_shadow            (bool)
      wm_opacity           (int 0..100)
      wm_margin            (int px)
      wm_image_path        (str)
      wm_image_size_pct    (int 1..50)
    """

    # current (1-based), total, current filename
    progress = pyqtSignal(int, int, str)
    # processed_count, error_count, errors (list of (filename, message))
    finished = pyqtSignal(int, int, list)

    def __init__(self, config: dict):
        super().__init__()
        self.cfg = config
        self._cancel = False

    def cancel(self):
        self._cancel = True

    # ── Run ───────────────────────────────────────────────────────────────
    def run(self):
        cfg = self.cfg
        in_dir = Path(cfg.get("input_dir", ""))
        overwrite = bool(cfg.get("overwrite"))
        out_dir = None if overwrite else Path(cfg.get("output_dir", ""))

        errors: list[tuple[str, str]] = []
        processed = 0

        try:
            files = sorted(
                p for p in in_dir.iterdir()
                if p.is_file() and p.suffix.lower() in SUPPORTED_EXT
            )
        except Exception as e:
            self.finished.emit(0, 1, [("<input dir>", str(e))])
            return

        total = len(files)
        if total == 0:
            self.finished.emit(0, 0, [])
            return

        if out_dir is not None:
            try:
                out_dir.mkdir(parents=True, exist_ok=True)
            except Exception as e:
                self.finished.emit(0, 1, [("<output dir>", str(e))])
                return

        for i, src in enumerate(files, start=1):
            if self._cancel:
                break
            self.progress.emit(i, total, src.name)
            try:
                dst = src if overwrite else (out_dir / src.name)
                self._process_one(src, dst)
                processed += 1
            except Exception as e:  # keep going on per-file errors
                errors.append((src.name, str(e)))

        self.finished.emit(processed, len(errors), errors)

    # ── Per-file pipeline ─────────────────────────────────────────────────
    def _process_one(self, src: Path, dst: Path):
        cfg = self.cfg
        img = PILImage.open(src)
        # Preserve original mode info for later save
        orig_format = (img.format or "JPEG").upper()
        if orig_format == "JPG":
            orig_format = "JPEG"

        # Convert for editing; keep alpha if the source had one
        if img.mode not in ("RGB", "RGBA"):
            img = img.convert("RGBA" if "A" in img.mode else "RGB")

        # ── Crop ──
        ratio = cfg.get("crop_ratio")
        if ratio:
            img = self._crop_center(img, ratio)

        # ── Resize ──
        img = self._apply_resize(img, cfg)

        # ── Watermark ──
        mode = cfg.get("watermark_mode", "none")
        if mode == "date":
            img = self._apply_date_watermark(img, cfg)
        elif mode == "image":
            img = self._apply_image_watermark(img, cfg)

        # ── Save ──
        save_kwargs = {}
        save_mode_img = img
        if orig_format == "PNG":
            save_kwargs["optimize"] = True

        # Normalize for JPEG: JPEG only supports RGB, no alpha / palette / CMYK / grayscale
        if orig_format in ("JPEG", "JPG"):
            if save_mode_img.mode == "RGBA":
                bg = PILImage.new("RGB", save_mode_img.size, (255, 255, 255))
                bg.paste(save_mode_img, mask=save_mode_img.split()[-1])
                save_mode_img = bg
            elif save_mode_img.mode != "RGB":
                save_mode_img = save_mode_img.convert("RGB")
            save_kwargs.setdefault("quality", 92)

        dst.parent.mkdir(parents=True, exist_ok=True)
        save_mode_img.save(dst, format=orig_format, **save_kwargs)

        # ── Metadata ──
        self._apply_metadata(dst, cfg)

    # ── Resize ────────────────────────────────────────────────────────────
    @staticmethod
    def _apply_resize(img, cfg):
        mode = cfg.get("resize_mode", "none")
        w, h = img.size
        if mode == "none" or (w == 0 or h == 0):
            return img
        if mode == "long_side":
            target = int(cfg.get("resize_long_side", 1920))
            long_side = max(w, h)
            if long_side <= 0:
                return img
            ratio = target / long_side
            new_w = max(1, int(round(w * ratio)))
            new_h = max(1, int(round(h * ratio)))
            return img.resize((new_w, new_h), PILImage.LANCZOS)
        if mode == "percent":
            pct = int(cfg.get("resize_percent", 100))
            if pct == 100:
                return img
            new_w = max(1, int(round(w * pct / 100)))
            new_h = max(1, int(round(h * pct / 100)))
            return img.resize((new_w, new_h), PILImage.LANCZOS)
        if mode == "exact":
            tw = int(cfg.get("resize_w", w))
            th = int(cfg.get("resize_h", h))
            if cfg.get("resize_keep_aspect", True):
                ratio = min(tw / w, th / h)
                new_w = max(1, int(round(w * ratio)))
                new_h = max(1, int(round(h * ratio)))
                return img.resize((new_w, new_h), PILImage.LANCZOS)
            return img.resize((max(1, tw), max(1, th)), PILImage.LANCZOS)
        return img

    # ── Crop ──────────────────────────────────────────────────────────────
    @staticmethod
    def _crop_center(img, ratio):
        w, h = img.size
        target_aspect = ratio[0] / ratio[1]
        current_aspect = w / h
        if current_aspect > target_aspect:
            new_w = int(h * target_aspect)
            left = (w - new_w) // 2
            return img.crop((left, 0, left + new_w, h))
        new_h = int(w / target_aspect)
        top = (h - new_h) // 2
        return img.crop((0, top, w, top + new_h))

    # ── Watermark: date ───────────────────────────────────────────────────
    def _apply_date_watermark(self, img, cfg):
        dt = cfg.get("wm_date_dt") or datetime.now()
        text = format_date(dt, int(cfg.get("wm_date_format_index", 0)))
        color = tuple(cfg.get("wm_color", (255, 255, 255)))
        size_pct = int(cfg.get("wm_font_size_pct", 5))
        font = self._resolve_font(cfg.get("wm_font", "Arial"),
                                  max(8, int(img.width * size_pct / 100)))
        position = cfg.get("wm_position", "Bottom-right")
        margin = int(cfg.get("wm_margin", 20))
        opacity = int(cfg.get("wm_opacity", 100))
        shadow = bool(cfg.get("wm_shadow", False))

        return self._render_text_watermark(
            img, text, font, color, position, margin, opacity, shadow,
        )

    @staticmethod
    def _resolve_font(family: str, size: int):
        """Try to load a TTF font; fall back to Pillow default."""
        # Common font filename guesses per family. Pillow searches the system
        # font path, so just-the-name usually works on Windows/macOS.
        candidates = [
            family,
            f"{family}.ttf",
            f"{family}.ttc",
            family.replace(" ", ""),
            f"{family.replace(' ', '')}.ttf",
        ]
        # OS-aware fallbacks
        sys_name = platform.system()
        if sys_name == "Windows":
            candidates += [
                "arial.ttf", "Arial.ttf",
                "C:/Windows/Fonts/arial.ttf",
            ]
        elif sys_name == "Darwin":
            candidates += [
                "/Library/Fonts/Arial.ttf",
                "/System/Library/Fonts/Helvetica.ttc",
                "/System/Library/Fonts/Supplemental/Arial.ttf",
            ]
        else:
            candidates += [
                "/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf",
                "/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf",
                "DejaVuSans.ttf",
            ]
        for c in candidates:
            try:
                return ImageFont.truetype(c, size)
            except Exception:
                continue
        # Last resort
        try:
            return ImageFont.load_default()
        except Exception:
            return None

    @staticmethod
    def _text_size(draw, text, font):
        """Return (w, h) for `text` using Pillow's bbox (works across versions)."""
        try:
            bbox = draw.textbbox((0, 0), text, font=font)
            return bbox[2] - bbox[0], bbox[3] - bbox[1]
        except Exception:
            try:
                return draw.textsize(text, font=font)
            except Exception:
                return (len(text) * 8, 14)

    @staticmethod
    def _anchor_xy(img_w, img_h, w, h, position, margin):
        pos = position.lower()
        if "bottom" in pos:
            y = img_h - h - margin
        elif "top" in pos:
            y = margin
        else:  # center
            y = (img_h - h) // 2
        if "right" in pos:
            x = img_w - w - margin
        elif "left" in pos:
            x = margin
        else:
            x = (img_w - w) // 2
        return x, y

    def _render_text_watermark(self, base, text, font, color, position,
                               margin, opacity, shadow):
        if base.mode != "RGBA":
            base = base.convert("RGBA")
        layer = PILImage.new("RGBA", base.size, (0, 0, 0, 0))
        draw = ImageDraw.Draw(layer)
        tw, th = self._text_size(draw, text, font)
        x, y = self._anchor_xy(base.width, base.height, tw, th, position, margin)
        alpha = max(0, min(255, int(opacity * 255 / 100)))
        rgba = (color[0], color[1], color[2], alpha)
        shadow_rgba = (0, 0, 0, alpha)
        if shadow:
            # 1px in 4 directions for a cheap outline
            for dx, dy in ((-1, 0), (1, 0), (0, -1), (0, 1)):
                draw.text((x + dx, y + dy), text, font=font, fill=shadow_rgba)
        draw.text((x, y), text, font=font, fill=rgba)
        return PILImage.alpha_composite(base, layer)

    # ── Watermark: image ──────────────────────────────────────────────────
    def _apply_image_watermark(self, img, cfg):
        wm_path = cfg.get("wm_image_path", "")
        if not wm_path or not os.path.exists(wm_path):
            return img
        try:
            wm = PILImage.open(wm_path).convert("RGBA")
        except Exception:
            return img
        size_pct = int(cfg.get("wm_image_size_pct", 15))
        target_w = max(1, int(img.width * size_pct / 100))
        ratio = target_w / wm.width
        target_h = max(1, int(wm.height * ratio))
        wm = wm.resize((target_w, target_h), PILImage.LANCZOS)

        opacity = int(cfg.get("wm_opacity", 100))
        if opacity < 100:
            r, g, b, a = wm.split()
            a = a.point(lambda v: int(v * opacity / 100))
            wm = PILImage.merge("RGBA", (r, g, b, a))

        position = cfg.get("wm_position", "Bottom-right")
        margin = int(cfg.get("wm_margin", 20))
        x, y = self._anchor_xy(img.width, img.height, wm.width, wm.height,
                               position, margin)

        if img.mode != "RGBA":
            img = img.convert("RGBA")
        layer = PILImage.new("RGBA", img.size, (0, 0, 0, 0))
        layer.paste(wm, (x, y), wm)
        return PILImage.alpha_composite(img, layer)

    # ── Metadata writing ──────────────────────────────────────────────────
    @staticmethod
    def _apply_metadata(path: Path, cfg: dict):
        created = cfg.get("created_dt")
        modified = cfg.get("modified_dt")
        write_exif = bool(cfg.get("write_exif"))

        # atime / mtime via os.utime — cross-platform.
        if modified is not None:
            try:
                ts = modified.timestamp()
                os.utime(path, (ts, ts))
            except Exception:
                pass

        # Created time: Windows-only via pywin32.
        if created is not None and platform.system() == "Windows":
            try:
                import pywintypes  # type: ignore
                import win32file   # type: ignore
                import win32con    # type: ignore
                handle = win32file.CreateFile(
                    str(path),
                    win32con.GENERIC_WRITE,
                    win32con.FILE_SHARE_READ | win32con.FILE_SHARE_WRITE,
                    None,
                    win32con.OPEN_EXISTING,
                    0,
                    None,
                )
                wt = pywintypes.Time(created.timestamp())
                win32file.SetFileTime(handle, wt, None, None)
                handle.close()
            except Exception:
                pass

        # EXIF DateTimeOriginal / DateTime via piexif (optional, JPEG/TIFF only).
        if write_exif and created is not None:
            ext = path.suffix.lower()
            if ext in (".jpg", ".jpeg", ".tif", ".tiff"):
                try:
                    import piexif  # type: ignore
                    stamp = created.strftime("%Y:%m:%d %H:%M:%S").encode("ascii")
                    try:
                        exif_dict = piexif.load(str(path))
                    except Exception:
                        exif_dict = {"0th": {}, "Exif": {}, "GPS": {}, "1st": {}, "thumbnail": None}
                    exif_dict.setdefault("0th", {})[piexif.ImageIFD.DateTime] = stamp
                    exif_dict.setdefault("Exif", {})[piexif.ExifIFD.DateTimeOriginal] = stamp
                    exif_dict["Exif"][piexif.ExifIFD.DateTimeDigitized] = stamp
                    piexif.insert(piexif.dump(exif_dict), str(path))
                except Exception:
                    pass
