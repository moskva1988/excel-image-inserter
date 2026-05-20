import os

from PyQt5.QtCore import QThread, pyqtSignal
from PIL import Image as PILImage


def estimate_size(path, max_w, max_h):
    """Return (original_mb, estimated_mb, width, height)."""
    size_mb = os.path.getsize(path) / (1024 * 1024)
    try:
        img = PILImage.open(path)
        w, h = img.size
        if max_w or max_h:
            ratio = 1.0
            if max_w and max_h:
                ratio = min(max_w / w, max_h / h)
            elif max_w:
                ratio = max_w / w
            else:
                ratio = max_h / h
            if ratio < 1:
                new_pixels = int(w * ratio) * int(h * ratio)
            else:
                new_pixels = w * h
            est_mb = new_pixels * 0.5 / (1024 * 1024)
        else:
            est_mb = size_mb
        return size_mb, est_mb, w, h
    except Exception:
        return size_mb, size_mb, 0, 0


# ── Image loader thread ───────────────────────────────────────────────────────
class ImageLoaderThread(QThread):
    progress = pyqtSignal(int, int)
    item_ready = pyqtSignal(str, float, float, int, int)
    finished = pyqtSignal()

    def __init__(self, paths, max_w, max_h):
        super().__init__()
        self.paths = paths
        self.max_w = max_w
        self.max_h = max_h

    def run(self):
        total = len(self.paths)
        for i, p in enumerate(self.paths):
            orig_mb, est_mb, w, h = estimate_size(p, self.max_w, self.max_h)
            self.item_ready.emit(p, orig_mb, est_mb, w, h)
            self.progress.emit(i + 1, total)
        self.finished.emit()
