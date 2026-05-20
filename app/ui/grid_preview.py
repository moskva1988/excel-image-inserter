import math

from PyQt5.QtWidgets import QWidget
from PyQt5.QtCore import Qt
from PyQt5.QtGui import QPainter, QPen, QColor, QFont
from openpyxl.utils import get_column_letter


# ── Grid preview widget ───────────────────────────────────────────────────────
class GridPreview(QWidget):
    HEADER_H = 16
    ROW_NUM_W = 24

    def __init__(self):
        super().__init__()
        self.groups = []
        self.cols = 2
        self.crop_ratio = None
        self.start_col = "A"
        self.start_row = 1
        self.placement = "over"
        self.use_groups = False
        self.setMinimumHeight(100)
        self.setMaximumHeight(160)

    def update_params(self, groups, cols, crop_ratio, start_col="A", start_row=1, placement="over", use_groups=False):
        self.groups = groups
        self.cols = max(1, cols)
        self.crop_ratio = crop_ratio
        self.start_col = start_col
        self.start_row = start_row
        self.placement = placement
        self.use_groups = use_groups
        self.update()

    def paintEvent(self, event):
        painter = QPainter(self)
        painter.setRenderHint(QPainter.Antialiasing)
        total_w = self.width()
        total_h = self.height()
        hh = self.HEADER_H
        rw = self.ROW_NUM_W
        painter.fillRect(self.rect(), QColor("#f0f0f0"))

        total_images = sum(len(g["images"]) for g in self.groups)
        if total_images == 0:
            self._draw_headers(painter, total_w, total_h, hh, rw, 3, 5)
            painter.end()
            return

        start_col_idx = self._col_to_idx(self.start_col)
        content_rows = 0
        for g in self.groups:
            if self.use_groups:
                content_rows += 1
            content_rows += math.ceil(len(g["images"]) / self.cols) if g["images"] else 0
            if self.use_groups:
                content_rows += 1

        show_cols = max(self.cols + start_col_idx, start_col_idx + self.cols + 1)
        show_rows = max(content_rows + self.start_row, self.start_row + content_rows + 1)
        self._draw_headers(painter, total_w, total_h, hh, rw, show_cols, show_rows)

        grid_w = total_w - rw
        grid_h = total_h - hh
        cell_w = grid_w / show_cols if show_cols else grid_w
        cell_h = grid_h / show_rows if show_rows else grid_h
        aspect = (self.crop_ratio[0] / self.crop_ratio[1]) if self.crop_ratio else 4 / 3

        current_row = self.start_row - 1
        img_num = 0

        for g in self.groups:
            if self.use_groups:
                hx = rw + start_col_idx * cell_w + 2
                hy = hh + current_row * cell_h
                painter.setPen(QColor("#1a1a1a"))
                painter.setFont(QFont("Arial", 7, QFont.Bold))
                painter.drawText(int(hx), int(hy), int(cell_w * self.cols), int(cell_h),
                                 Qt.AlignLeft | Qt.AlignVCenter, g["title"])
                current_row += 1

            img_rows = math.ceil(len(g["images"]) / self.cols) if g["images"] else 0
            for r in range(img_rows):
                for c in range(self.cols):
                    idx = r * self.cols + c
                    if idx >= len(g["images"]):
                        break
                    img_num += 1
                    grid_col = start_col_idx + c
                    grid_row = current_row + r
                    cx = rw + grid_col * cell_w + 1
                    cy = hh + grid_row * cell_h + 1
                    cw = cell_w - 2
                    ch = cell_h - 2
                    img_aspect = aspect
                    cell_aspect = cw / ch if ch > 0 else 1
                    if img_aspect > cell_aspect:
                        iw, ih = cw, cw / img_aspect
                    else:
                        ih, iw = ch, ch * img_aspect
                    ix = cx + (cw - iw) / 2
                    iy = cy + (ch - ih) / 2
                    painter.fillRect(int(ix), int(iy), int(iw), int(ih), QColor("#6366f1"))
                    painter.setPen(QColor("#fff"))
                    painter.setFont(QFont("Arial", 7))
                    painter.drawText(int(ix), int(iy), int(iw), int(ih), Qt.AlignCenter, str(img_num))

            current_row += img_rows + (1 if self.use_groups else 0)

        painter.end()

    def _draw_headers(self, painter, total_w, total_h, hh, rw, show_cols, show_rows):
        grid_w = total_w - rw
        grid_h = total_h - hh
        cell_w = grid_w / show_cols if show_cols else grid_w
        cell_h = grid_h / show_rows if show_rows else grid_h
        painter.fillRect(rw, 0, int(grid_w), hh, QColor("#e0e0e0"))
        painter.fillRect(0, hh, rw, int(grid_h), QColor("#e0e0e0"))
        painter.fillRect(0, 0, rw, hh, QColor("#d0d0d0"))
        painter.setPen(QPen(QColor("#c0c0c0"), 1))
        for c in range(show_cols + 1):
            x = int(rw + c * cell_w)
            painter.drawLine(x, 0, x, total_h)
        for r in range(show_rows + 1):
            y = int(hh + r * cell_h)
            painter.drawLine(0, y, total_w, y)
        painter.setPen(QPen(QColor("#999"), 1))
        painter.drawLine(0, hh, total_w, hh)
        painter.drawLine(rw, 0, rw, total_h)
        painter.setPen(QColor("#333"))
        painter.setFont(QFont("Arial", 7))
        for c in range(show_cols):
            x = int(rw + c * cell_w)
            letter = get_column_letter(c + 1)
            painter.drawText(x, 0, int(cell_w), hh, Qt.AlignCenter, letter)
        for r in range(show_rows):
            y = int(hh + r * cell_h)
            painter.drawText(0, y, rw, int(cell_h), Qt.AlignCenter, str(r + 1))

    @staticmethod
    def _col_to_idx(col_str):
        col_str = col_str.upper().strip()
        idx = 0
        for ch in col_str:
            if ch.isalpha():
                idx = idx * 26 + (ord(ch) - ord('A'))
        return idx
