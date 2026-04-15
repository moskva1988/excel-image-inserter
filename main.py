#!/usr/bin/env python3
"""Excel Image Inserter — PyQt5 utility for batch-inserting images into Excel."""

import sys
import os
import math
from pathlib import Path
from io import BytesIO

from PyQt5.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
    QGroupBox, QLabel, QPushButton, QComboBox, QSpinBox, QDoubleSpinBox,
    QLineEdit, QFileDialog, QListWidget, QListWidgetItem, QAbstractItemView,
    QRadioButton, QButtonGroup, QMessageBox, QProgressBar, QCheckBox,
    QFrame, QGridLayout, QSizePolicy, QScrollArea, QToolTip,
    QTreeWidget, QTreeWidgetItem, QHeaderView,
)
from PyQt5.QtCore import Qt, QSize, QThread, pyqtSignal, QRect, QPoint
from PyQt5.QtGui import QPixmap, QIcon, QImage, QPainter, QPen, QColor, QFont, QBrush

from PIL import Image as PILImage
import openpyxl
from openpyxl.drawing.image import Image as XLImage
from openpyxl.drawing.spreadsheet_drawing import OneCellAnchor, AnchorMarker
from openpyxl.drawing.xdr import XDRPositiveSize2D
from openpyxl.utils import get_column_letter
from openpyxl.utils.units import pixels_to_EMU


# ── Constants ──────────────────────────────────────────────────────────────────
CM_TO_PX_96 = 96 / 2.54
CROP_PRESETS = {
    "None": None,
    "1:1": (1, 1),
    "4:3": (4, 3),
    "3:2": (3, 2),
    "16:9": (16, 9),
    "3:4": (3, 4),
    "2:3": (2, 3),
    "9:16": (9, 16),
}
THUMB_SIZE = 100


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


# ── Worker thread ──────────────────────────────────────────────────────────────
class InsertWorker(QThread):
    progress = pyqtSignal(int)
    finished = pyqtSignal(str)
    status = pyqtSignal(str)

    def __init__(self, params):
        super().__init__()
        self.p = params

    def run(self):
        try:
            self._do_insert()
            self.finished.emit("")
        except Exception as e:
            self.finished.emit(str(e))

    def _do_insert(self):
        p = self.p
        if p["excel_path"] and os.path.exists(p["excel_path"]):
            wb = openpyxl.load_workbook(p["excel_path"])
        else:
            wb = openpyxl.Workbook()

        if p["sheet_new"]:
            ws = wb.create_sheet(title=p["sheet_name"])
        else:
            ws = wb[p["sheet_name"]]

        total = len(p["images"])
        cols = p["grid_cols"]
        start_col_idx = openpyxl.utils.column_index_from_string(p["start_col"])
        start_row = p["start_row"]

        for i, img_path in enumerate(p["images"]):
            self.status.emit(f"Processing {i+1}/{total}: {Path(img_path).name}")

            img = PILImage.open(img_path).convert("RGB")

            if p["crop_ratio"]:
                img = self._crop_center(img, p["crop_ratio"])
            if p["resize_px_w"] or p["resize_px_h"]:
                img = self._resize_px(img, p["resize_px_w"], p["resize_px_h"])

            buf = BytesIO()
            img.save(buf, format="JPEG", quality=90)
            buf.seek(0)

            xl_img = XLImage(buf)
            w_cm = p["display_w_cm"]
            h_cm = p["display_h_cm"]
            xl_img.width = w_cm * CM_TO_PX_96
            xl_img.height = h_cm * CM_TO_PX_96

            row_offset = i // cols
            col_offset = i % cols
            img_w_px = xl_img.width
            img_h_px = xl_img.height

            if p["placement"] == "in_cell":
                cell_col = start_col_idx + col_offset
                cell_row = start_row + row_offset
                ws.column_dimensions[get_column_letter(cell_col)].width = w_cm * 4.8
                ws.row_dimensions[cell_row].height = h_cm * 28.35
                ws.add_image(xl_img, f"{get_column_letter(cell_col)}{cell_row}")
            else:
                # "Over cells" — place images using pixel offsets, never resize cells
                gap_px = 10
                x_px = int(col_offset * (img_w_px + gap_px))
                y_px = int(row_offset * (img_h_px + gap_px))
                emu_w = pixels_to_EMU(img_w_px)
                emu_h = pixels_to_EMU(img_h_px)
                # Calculate which cell + offset for x
                col_i = start_col_idx - 1  # 0-based
                remaining_x = x_px
                # Default Excel column width ~64px, row height ~20px
                default_col_px = 64
                default_row_px = 20
                while remaining_x > default_col_px:
                    remaining_x -= default_col_px
                    col_i += 1
                row_i = start_row - 1  # 0-based
                remaining_y = y_px
                while remaining_y > default_row_px:
                    remaining_y -= default_row_px
                    row_i += 1
                marker = AnchorMarker(
                    col=col_i, colOff=pixels_to_EMU(remaining_x),
                    row=row_i, rowOff=pixels_to_EMU(remaining_y),
                )
                anchor = OneCellAnchor(
                    _from=marker,
                    ext=XDRPositiveSize2D(cx=emu_w, cy=emu_h),
                )
                xl_img.anchor = anchor
                ws.add_image(xl_img)

            self.progress.emit(int((i + 1) / total * 100))

        save_path = p["save_path"] or p["excel_path"]
        self.status.emit(f"Saving {save_path}...")
        wb.save(save_path)

    @staticmethod
    def _crop_center(img, ratio):
        w, h = img.size
        target_aspect = ratio[0] / ratio[1]
        current_aspect = w / h
        if current_aspect > target_aspect:
            new_w = int(h * target_aspect)
            left = (w - new_w) // 2
            return img.crop((left, 0, left + new_w, h))
        else:
            new_h = int(w / target_aspect)
            top = (h - new_h) // 2
            return img.crop((0, top, w, top + new_h))

    @staticmethod
    def _resize_px(img, max_w, max_h):
        w, h = img.size
        if max_w and max_h:
            ratio = min(max_w / w, max_h / h)
        elif max_w:
            ratio = max_w / w
        else:
            ratio = max_h / h
        if ratio < 1:
            img = img.resize((int(w * ratio), int(h * ratio)), PILImage.LANCZOS)
        return img


# ── Thumbnail stack widget ─────────────────────────────────────────────────────
class ThumbCard(QWidget):
    """Single image card with overlay info and delete button."""
    delete_requested = pyqtSignal(str)  # path
    selection_toggled = pyqtSignal(str, bool)  # path, selected

    def __init__(self, path, orig_mb, est_mb, w, h):
        super().__init__()
        self.path = path
        self.orig_mb = orig_mb
        self.est_mb = est_mb
        self.img_w = w
        self.img_h = h
        self.selected = False
        self.setFixedSize(THUMB_SIZE + 10, THUMB_SIZE + 10)
        self.setToolTip(f"{Path(path).name}\n{w}x{h}\n{orig_mb:.1f} MB → {est_mb:.1f} MB")

        self.pixmap = QPixmap(path).scaled(
            THUMB_SIZE, THUMB_SIZE, Qt.KeepAspectRatio, Qt.SmoothTransformation
        )

    def paintEvent(self, event):
        p = QPainter(self)
        p.setRenderHint(QPainter.Antialiasing)

        # Background
        if self.selected:
            p.fillRect(self.rect(), QColor("#2a2a4a"))
            p.setPen(QPen(QColor("#6366f1"), 2))
            p.drawRect(1, 1, self.width() - 2, self.height() - 2)
        else:
            p.fillRect(self.rect(), QColor("#1a1a2e"))

        # Image centered
        x = (self.width() - self.pixmap.width()) // 2
        y = (self.height() - self.pixmap.height()) // 2
        p.drawPixmap(x, y, self.pixmap)

        # Bottom overlay bar
        bar_h = 16
        bar_y = self.height() - bar_h - 3
        p.fillRect(3, bar_y, self.width() - 6, bar_h, QColor(0, 0, 0, 160))

        p.setFont(QFont("Arial", 7))
        # Original size — bottom left
        p.setPen(QColor("#ccc"))
        p.drawText(6, bar_y, self.width() // 2, bar_h, Qt.AlignLeft | Qt.AlignVCenter,
                   f"{self.orig_mb:.1f}MB")
        # Estimated size — bottom right
        p.setPen(QColor("#22c55e"))
        p.drawText(self.width() // 2, bar_y, self.width() // 2 - 6, bar_h,
                   Qt.AlignRight | Qt.AlignVCenter, f"{self.est_mb:.1f}MB")

        # Delete button — top right
        btn_size = 16
        bx = self.width() - btn_size - 3
        by = 3
        p.fillRect(bx, by, btn_size, btn_size, QColor(200, 0, 0, 180))
        p.setPen(QColor("#fff"))
        p.setFont(QFont("Arial", 9, QFont.Bold))
        p.drawText(bx, by, btn_size, btn_size, Qt.AlignCenter, "×")

        p.end()

    def mousePressEvent(self, event):
        # Check if delete button clicked
        btn_size = 16
        bx = self.width() - btn_size - 3
        by = 3
        if QRect(bx, by, btn_size, btn_size).contains(event.pos()):
            self.delete_requested.emit(self.path)
            return
        # Toggle selection
        self.selected = not self.selected
        self.selection_toggled.emit(self.path, self.selected)
        self.update()


class ThumbStackView(QScrollArea):
    """Flow layout of ThumbCard widgets."""
    delete_requested = pyqtSignal(str)

    def __init__(self):
        super().__init__()
        self.setWidgetResizable(True)
        self.container = QWidget()
        self.flow = FlowLayout(self.container)
        self.flow.setSpacing(6)
        self.setWidget(self.container)
        self.cards = []
        self.selected_paths = set()

    def set_images(self, paths, max_w, max_h):
        # Clear
        for c in self.cards:
            c.deleteLater()
        self.cards.clear()
        self.selected_paths.clear()

        for path in paths:
            orig_mb, est_mb, w, h = estimate_size(path, max_w, max_h)
            card = ThumbCard(path, orig_mb, est_mb, w, h)
            card.delete_requested.connect(self._on_delete)
            card.selection_toggled.connect(self._on_selection)
            self.flow.addWidget(card)
            self.cards.append(card)

    def _on_delete(self, path):
        self.delete_requested.emit(path)

    def _on_selection(self, path, selected):
        if selected:
            self.selected_paths.add(path)
        else:
            self.selected_paths.discard(path)

    def get_selected(self):
        return list(self.selected_paths)


class FlowLayout(QVBoxLayout):
    """Simple flow layout using nested HBoxLayouts."""
    def __init__(self, parent=None):
        super().__init__(parent)
        self._widgets = []

    def addWidget(self, widget):
        self._widgets.append(widget)
        self._relayout()

    def _relayout(self):
        # Remove old layouts
        while self.count():
            item = self.takeAt(0)
            if item.layout():
                while item.layout().count():
                    item.layout().takeAt(0)

        if not self._widgets:
            return

        parent = self.parentWidget()
        available_w = parent.width() if parent else 600
        card_w = THUMB_SIZE + 16

        cols = max(1, available_w // card_w)
        row_lay = None
        for i, w in enumerate(self._widgets):
            if i % cols == 0:
                row_lay = QHBoxLayout()
                row_lay.setAlignment(Qt.AlignLeft)
                super().addLayout(row_lay)
            row_lay.addWidget(w)


# ── Grid preview widget ───────────────────────────────────────────────────────
class GridPreview(QWidget):
    def __init__(self):
        super().__init__()
        self.cols = 2
        self.rows = 0
        self.count = 0
        self.crop_ratio = None
        self.setMinimumHeight(80)
        self.setMaximumHeight(120)

    def update_params(self, cols, rows, count, crop_ratio):
        self.cols = max(1, cols)
        self.rows = rows
        self.count = count
        self.crop_ratio = crop_ratio
        self.update()

    def paintEvent(self, event):
        painter = QPainter(self)
        painter.setRenderHint(QPainter.Antialiasing)

        w = self.width() - 10
        h = self.height() - 10
        ox, oy = 5, 5

        painter.fillRect(self.rect(), QColor("#1a1a2e"))
        painter.setPen(QPen(QColor("#333"), 1))
        painter.drawRect(ox, oy, w, h)

        if self.count == 0:
            painter.end()
            return

        cols = self.cols
        actual_rows = math.ceil(self.count / cols) if cols else 1
        if self.rows > 0:
            actual_rows = min(actual_rows, self.rows)

        cell_w = w / cols
        cell_h = h / actual_rows
        gap = 2

        aspect = (self.crop_ratio[0] / self.crop_ratio[1]) if self.crop_ratio else 4 / 3

        drawn = 0
        for r in range(actual_rows):
            for c in range(cols):
                if drawn >= self.count:
                    break
                cx = ox + c * cell_w + gap
                cy = oy + r * cell_h + gap
                cw = cell_w - gap * 2
                ch = cell_h - gap * 2

                img_aspect = aspect
                cell_aspect = cw / ch if ch else 1
                if img_aspect > cell_aspect:
                    iw, ih = cw, cw / img_aspect
                else:
                    ih, iw = ch, ch * img_aspect
                ix = cx + (cw - iw) / 2
                iy = cy + (ch - ih) / 2

                painter.setPen(QPen(QColor("#444"), 1, Qt.DashLine))
                painter.drawRect(int(cx), int(cy), int(cw), int(ch))
                painter.fillRect(int(ix), int(iy), int(iw), int(ih), QColor("#6366f1"))
                painter.setPen(QColor("#fff"))
                painter.setFont(QFont("Arial", 7))
                painter.drawText(int(ix), int(iy), int(iw), int(ih), Qt.AlignCenter, str(drawn + 1))
                drawn += 1

        painter.end()


# ── Main window ────────────────────────────────────────────────────────────────
class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Excel Image Inserter")
        self.setMinimumSize(750, 750)
        self.image_paths = []
        self._build_ui()

    def _build_ui(self):
        central = QWidget()
        self.setCentralWidget(central)
        root = QVBoxLayout(central)
        root.setSpacing(6)

        # ── Excel file ─────────────────────────────────────────────────────
        grp_file = QGroupBox("Excel File")
        lay_file = QVBoxLayout(grp_file)
        lay_file.setSpacing(4)

        row1 = QHBoxLayout()
        self.rb_new = QRadioButton("Create new")
        self.rb_open = QRadioButton("Open existing")
        self.rb_open.setChecked(True)
        bg = QButtonGroup(self)
        bg.addButton(self.rb_new)
        bg.addButton(self.rb_open)
        row1.addWidget(self.rb_open)
        row1.addWidget(self.rb_new)
        lay_file.addLayout(row1)

        row2 = QHBoxLayout()
        self.le_file = QLineEdit()
        self.le_file.setPlaceholderText("Path to .xlsx file")
        self.btn_browse_file = QPushButton("Browse...")
        self.btn_browse_file.clicked.connect(self._browse_file)
        row2.addWidget(self.le_file, 1)
        row2.addWidget(self.btn_browse_file)
        lay_file.addLayout(row2)

        row3 = QHBoxLayout()
        row3.addWidget(QLabel("Sheet:"))
        self.combo_sheet = QComboBox()
        self.combo_sheet.setMinimumWidth(120)
        row3.addWidget(self.combo_sheet, 1)
        self.cb_new_sheet = QCheckBox("New:")
        self.le_new_sheet = QLineEdit()
        self.le_new_sheet.setPlaceholderText("Sheet name")
        self.le_new_sheet.setEnabled(False)
        self.cb_new_sheet.toggled.connect(self.le_new_sheet.setEnabled)
        self.cb_new_sheet.toggled.connect(lambda v: self.combo_sheet.setEnabled(not v))
        row3.addWidget(self.cb_new_sheet)
        row3.addWidget(self.le_new_sheet)
        lay_file.addLayout(row3)

        self.rb_new.toggled.connect(self._on_file_mode_changed)
        root.addWidget(grp_file)

        # ── Images ─────────────────────────────────────────────────────────
        grp_img = QGroupBox("Images")
        lay_img = QVBoxLayout(grp_img)

        btn_row = QHBoxLayout()
        self.btn_add_img = QPushButton("Add...")
        self.btn_add_img.clicked.connect(self._add_images)
        self.btn_remove_img = QPushButton("Remove")
        self.btn_remove_img.clicked.connect(self._remove_selected)
        self.btn_clear_img = QPushButton("Clear")
        self.btn_clear_img.clicked.connect(self._clear_images)
        btn_row.addWidget(self.btn_add_img)
        btn_row.addWidget(self.btn_remove_img)
        btn_row.addWidget(self.btn_clear_img)
        btn_row.addStretch()

        self.btn_view_thumb_list = QPushButton("List")
        self.btn_view_detail = QPushButton("Details")
        self.btn_view_thumb_stack = QPushButton("Stack")
        for b in [self.btn_view_thumb_list, self.btn_view_detail, self.btn_view_thumb_stack]:
            b.setCheckable(True)
            b.setMaximumWidth(65)
            b.setStyleSheet("QPushButton:checked{background:#6366f1;color:#fff;border-radius:4px;padding:3px 6px}")
        self.btn_view_thumb_list.setChecked(True)
        self.btn_view_thumb_list.clicked.connect(lambda: self._switch_view("list"))
        self.btn_view_detail.clicked.connect(lambda: self._switch_view("detail"))
        self.btn_view_thumb_stack.clicked.connect(lambda: self._switch_view("stack"))
        btn_row.addWidget(self.btn_view_thumb_list)
        btn_row.addWidget(self.btn_view_detail)
        btn_row.addWidget(self.btn_view_thumb_stack)
        lay_img.addLayout(btn_row)

        # View: Thumbnail list  [icon] | name | size | est size | [x]
        self.tree_list = QTreeWidget()
        self.tree_list.setHeaderLabels(["", "File", "Size", "After", ""])
        self.tree_list.setIconSize(QSize(48, 48))
        self.tree_list.setSelectionMode(QAbstractItemView.ExtendedSelection)
        self.tree_list.setRootIsDecorated(False)
        self.tree_list.setColumnWidth(0, 56)
        self.tree_list.setColumnWidth(1, 200)
        self.tree_list.setColumnWidth(2, 70)
        self.tree_list.setColumnWidth(3, 70)
        self.tree_list.setColumnWidth(4, 30)
        self.tree_list.header().setStretchLastSection(False)
        self.tree_list.header().setSectionResizeMode(1, QHeaderView.Stretch)
        self.tree_list.setMinimumHeight(150)
        self.tree_list.itemClicked.connect(self._on_tree_click)
        lay_img.addWidget(self.tree_list)

        # View: Detail list (no thumbnails)
        self.tree_detail = QTreeWidget()
        self.tree_detail.setHeaderLabels(["File", "Dimensions", "Size", "After", ""])
        self.tree_detail.setSelectionMode(QAbstractItemView.ExtendedSelection)
        self.tree_detail.setRootIsDecorated(False)
        self.tree_detail.setColumnWidth(0, 220)
        self.tree_detail.setColumnWidth(1, 90)
        self.tree_detail.setColumnWidth(2, 70)
        self.tree_detail.setColumnWidth(3, 70)
        self.tree_detail.setColumnWidth(4, 30)
        self.tree_detail.header().setStretchLastSection(False)
        self.tree_detail.header().setSectionResizeMode(0, QHeaderView.Stretch)
        self.tree_detail.setMinimumHeight(150)
        self.tree_detail.itemClicked.connect(self._on_tree_detail_click)
        self.tree_detail.hide()
        lay_img.addWidget(self.tree_detail)

        # View: Thumbnail stack
        self.thumb_stack = ThumbStackView()
        self.thumb_stack.delete_requested.connect(self._delete_by_path)
        self.thumb_stack.setMinimumHeight(150)
        self.thumb_stack.hide()
        lay_img.addWidget(self.thumb_stack)

        # Bottom: count + total
        bottom_row = QHBoxLayout()
        self.lbl_img_count = QLabel("0 images")
        self.lbl_total_size = QLabel("")
        self.lbl_total_size.setStyleSheet("color:#999")
        bottom_row.addWidget(self.lbl_img_count)
        bottom_row.addStretch()
        bottom_row.addWidget(self.lbl_total_size)
        lay_img.addLayout(bottom_row)

        root.addWidget(grp_img, 1)  # stretch

        # ── Settings: Resize + Display in one row ──────────────────────────
        settings_row = QHBoxLayout()

        grp_resize = QGroupBox("Resize (px)")
        g_r = QGridLayout(grp_resize)
        g_r.setSpacing(4)
        g_r.addWidget(QLabel("W:"), 0, 0)
        self.spin_px_w = QSpinBox()
        self.spin_px_w.setRange(0, 10000)
        self.spin_px_w.setValue(1200)
        self.spin_px_w.setSpecialValueText("Auto")
        g_r.addWidget(self.spin_px_w, 0, 1)
        g_r.addWidget(QLabel("H:"), 1, 0)
        self.spin_px_h = QSpinBox()
        self.spin_px_h.setRange(0, 10000)
        self.spin_px_h.setValue(0)
        self.spin_px_h.setSpecialValueText("Auto")
        g_r.addWidget(self.spin_px_h, 1, 1)
        settings_row.addWidget(grp_resize)

        grp_display = QGroupBox("Display (cm)")
        g_d = QGridLayout(grp_display)
        g_d.setSpacing(4)
        g_d.addWidget(QLabel("W:"), 0, 0)
        self.spin_cm_w = QDoubleSpinBox()
        self.spin_cm_w.setRange(0.5, 50)
        self.spin_cm_w.setValue(6.0)
        self.spin_cm_w.setSingleStep(0.5)
        self.spin_cm_w.setSuffix(" cm")
        g_d.addWidget(self.spin_cm_w, 0, 1)
        g_d.addWidget(QLabel("H:"), 1, 0)
        self.spin_cm_h = QDoubleSpinBox()
        self.spin_cm_h.setRange(0.5, 50)
        self.spin_cm_h.setValue(4.5)
        self.spin_cm_h.setSingleStep(0.5)
        self.spin_cm_h.setSuffix(" cm")
        g_d.addWidget(self.spin_cm_h, 1, 1)
        settings_row.addWidget(grp_display)

        # Crop
        grp_crop = QGroupBox("Crop")
        g_c = QVBoxLayout(grp_crop)
        self.combo_crop = QComboBox()
        self.combo_crop.addItems(CROP_PRESETS.keys())
        self.combo_crop.currentTextChanged.connect(self._on_settings_changed)
        g_c.addWidget(self.combo_crop)
        g_c.addStretch()
        settings_row.addWidget(grp_crop)

        root.addLayout(settings_row)

        # ── Grid + Position + Preview in one row ───────────────────────────
        grid_row = QHBoxLayout()

        grp_grid = QGroupBox("Grid")
        g_g = QGridLayout(grp_grid)
        g_g.setSpacing(4)
        g_g.addWidget(QLabel("Cols:"), 0, 0)
        self.spin_cols = QSpinBox()
        self.spin_cols.setRange(1, 20)
        self.spin_cols.setValue(2)
        self.spin_cols.valueChanged.connect(self._on_settings_changed)
        g_g.addWidget(self.spin_cols, 0, 1)
        g_g.addWidget(QLabel("Rows:"), 1, 0)
        self.spin_rows = QSpinBox()
        self.spin_rows.setRange(0, 1000)
        self.spin_rows.setValue(0)
        self.spin_rows.setSpecialValueText("Auto")
        self.spin_rows.valueChanged.connect(self._on_settings_changed)
        g_g.addWidget(self.spin_rows, 1, 1)
        grid_row.addWidget(grp_grid)

        grp_pos = QGroupBox("Position")
        g_p = QGridLayout(grp_pos)
        g_p.setSpacing(4)
        g_p.addWidget(QLabel("Cell:"), 0, 0)
        pos_row = QHBoxLayout()
        self.le_start_col = QLineEdit("A")
        self.le_start_col.setMaximumWidth(35)
        self.spin_start_row = QSpinBox()
        self.spin_start_row.setRange(1, 1048576)
        self.spin_start_row.setValue(1)
        pos_row.addWidget(self.le_start_col)
        pos_row.addWidget(self.spin_start_row)
        g_p.addLayout(pos_row, 0, 1)
        g_p.addWidget(QLabel("Mode:"), 1, 0)
        self.combo_placement = QComboBox()
        self.combo_placement.addItems(["Over cells", "In cell"])
        g_p.addWidget(self.combo_placement, 1, 1)
        grid_row.addWidget(grp_pos)

        # Grid preview (compact, aligned with Crop top)
        self.grid_preview = GridPreview()
        grid_row.addWidget(self.grid_preview, 1)

        root.addLayout(grid_row)

        # ── Action ─────────────────────────────────────────────────────────
        self.progress = QProgressBar()
        self.progress.setValue(0)
        self.progress.setMaximumHeight(16)
        root.addWidget(self.progress)

        action_row = QHBoxLayout()
        self.lbl_status = QLabel("Ready")
        action_row.addWidget(self.lbl_status, 1)
        self.btn_insert = QPushButton("  Insert Images  ")
        self.btn_insert.setMinimumHeight(36)
        self.btn_insert.setStyleSheet("font-size:13px;font-weight:bold;background:#6366f1;color:#fff;border-radius:6px;padding:6px 20px")
        self.btn_insert.clicked.connect(self._do_insert)
        action_row.addWidget(self.btn_insert)
        root.addLayout(action_row)

    # ── Slots ──────────────────────────────────────────────────────────────
    def _on_file_mode_changed(self):
        is_open = self.rb_open.isChecked()
        self.btn_browse_file.setText("Browse..." if is_open else "Save as...")
        if not is_open:
            self.combo_sheet.clear()
            self.cb_new_sheet.setChecked(True)

    def _browse_file(self):
        if self.rb_open.isChecked():
            path, _ = QFileDialog.getOpenFileName(self, "Open Excel", "", "Excel Files (*.xlsx)")
        else:
            path, _ = QFileDialog.getSaveFileName(self, "Save Excel As", "images.xlsx", "Excel Files (*.xlsx)")
        if path:
            self.le_file.setText(path)
            if self.rb_open.isChecked() and os.path.exists(path):
                try:
                    wb = openpyxl.load_workbook(path, read_only=True)
                    self.combo_sheet.clear()
                    self.combo_sheet.addItems(wb.sheetnames)
                    wb.close()
                except Exception as e:
                    QMessageBox.warning(self, "Error", str(e))

    def _add_images(self):
        paths, _ = QFileDialog.getOpenFileNames(
            self, "Select Images", "",
            "Images (*.jpg *.jpeg *.png *.bmp *.webp *.tiff);;All Files (*)"
        )
        for p in paths:
            if p not in self.image_paths:
                self.image_paths.append(p)
        self._rebuild_views()

    def _clear_images(self):
        if self.image_paths:
            if QMessageBox.question(self, "Clear", f"Remove all {len(self.image_paths)} images?",
                                    QMessageBox.Yes | QMessageBox.No) != QMessageBox.Yes:
                return
        self.image_paths.clear()
        self._rebuild_views()

    def _remove_selected(self):
        # Collect selected paths from active view
        to_remove = set()

        if self.tree_list.isVisible():
            for item in self.tree_list.selectedItems():
                to_remove.add(item.data(0, Qt.UserRole))
        elif self.tree_detail.isVisible():
            for item in self.tree_detail.selectedItems():
                to_remove.add(item.data(0, Qt.UserRole))
        elif self.thumb_stack.isVisible():
            to_remove = set(self.thumb_stack.get_selected())

        if not to_remove:
            return

        if QMessageBox.question(self, "Remove",
                                f"Remove {len(to_remove)} image(s)?",
                                QMessageBox.Yes | QMessageBox.No) != QMessageBox.Yes:
            return

        self.image_paths = [p for p in self.image_paths if p not in to_remove]
        self._rebuild_views()

    def _delete_by_path(self, path):
        if QMessageBox.question(self, "Remove",
                                f"Remove {Path(path).name}?",
                                QMessageBox.Yes | QMessageBox.No) != QMessageBox.Yes:
            return
        if path in self.image_paths:
            self.image_paths.remove(path)
        self._rebuild_views()

    def _rebuild_views(self):
        max_w = self.spin_px_w.value() or None
        max_h = self.spin_px_h.value() or None

        # Thumbnail list
        self.tree_list.clear()
        for p in self.image_paths:
            orig_mb, est_mb, w, h = estimate_size(p, max_w, max_h)
            item = QTreeWidgetItem(["", Path(p).name, f"{orig_mb:.1f} MB", f"{est_mb:.1f} MB", "×"])
            item.setData(0, Qt.UserRole, p)
            try:
                px = QPixmap(p).scaled(48, 48, Qt.KeepAspectRatio, Qt.SmoothTransformation)
                item.setIcon(0, QIcon(px))
            except Exception:
                pass
            self.tree_list.addTopLevelItem(item)

        # Detail list
        self.tree_detail.clear()
        for p in self.image_paths:
            orig_mb, est_mb, w, h = estimate_size(p, max_w, max_h)
            dim = f"{w}×{h}" if w else "?"
            item = QTreeWidgetItem([Path(p).name, dim, f"{orig_mb:.1f} MB", f"{est_mb:.1f} MB", "×"])
            item.setData(0, Qt.UserRole, p)
            self.tree_detail.addTopLevelItem(item)

        # Thumb stack
        self.thumb_stack.set_images(self.image_paths, max_w, max_h)

        self._update_count()

    def _on_tree_click(self, item, col):
        if col == 4:  # delete column
            path = item.data(0, Qt.UserRole)
            if path:
                self._delete_by_path(path)

    def _on_tree_detail_click(self, item, col):
        if col == 4:
            path = item.data(0, Qt.UserRole)
            if path:
                self._delete_by_path(path)

    def _switch_view(self, mode):
        self.btn_view_thumb_list.setChecked(mode == "list")
        self.btn_view_detail.setChecked(mode == "detail")
        self.btn_view_thumb_stack.setChecked(mode == "stack")
        self.tree_list.setVisible(mode == "list")
        self.tree_detail.setVisible(mode == "detail")
        self.thumb_stack.setVisible(mode == "stack")

    def _update_count(self):
        n = len(self.image_paths)
        self.lbl_img_count.setText(f"{n} image{'s' if n != 1 else ''}")
        total_orig = sum(os.path.getsize(p) / (1024 * 1024) for p in self.image_paths if os.path.exists(p))
        max_w = self.spin_px_w.value() or None
        max_h = self.spin_px_h.value() or None
        total_est = sum(estimate_size(p, max_w, max_h)[1] for p in self.image_paths)
        if total_orig > 0:
            self.lbl_total_size.setText(f"Total: {total_orig:.1f} MB → {total_est:.1f} MB")
        else:
            self.lbl_total_size.setText("")
        self._on_settings_changed()

    def _on_settings_changed(self, *_):
        crop_key = self.combo_crop.currentText()
        crop = CROP_PRESETS.get(crop_key)
        self.grid_preview.update_params(
            cols=self.spin_cols.value(),
            rows=self.spin_rows.value(),
            count=len(self.image_paths),
            crop_ratio=crop,
        )

    def _do_insert(self):
        file_path = self.le_file.text().strip()
        if self.rb_open.isChecked() and (not file_path or not os.path.exists(file_path)):
            QMessageBox.warning(self, "Error", "Please select an existing Excel file.")
            return
        if self.rb_new.isChecked() and not file_path:
            QMessageBox.warning(self, "Error", "Please specify a file path to save.")
            return
        if not self.image_paths:
            QMessageBox.warning(self, "Error", "No images selected.")
            return

        sheet_new = self.cb_new_sheet.isChecked()
        if sheet_new:
            sheet_name = self.le_new_sheet.text().strip()
            if not sheet_name:
                QMessageBox.warning(self, "Error", "Enter a sheet name.")
                return
        else:
            sheet_name = self.combo_sheet.currentText()
            if not sheet_name:
                QMessageBox.warning(self, "Error", "Select a sheet.")
                return

        start_col = self.le_start_col.text().strip().upper()
        if not start_col or not start_col.isalpha():
            QMessageBox.warning(self, "Error", "Column must be a letter (A, B, C...).")
            return

        crop = CROP_PRESETS.get(self.combo_crop.currentText())

        params = {
            "excel_path": file_path if self.rb_open.isChecked() else None,
            "save_path": file_path,
            "sheet_new": sheet_new,
            "sheet_name": sheet_name,
            "images": list(self.image_paths),
            "resize_px_w": self.spin_px_w.value() or None,
            "resize_px_h": self.spin_px_h.value() or None,
            "display_w_cm": self.spin_cm_w.value(),
            "display_h_cm": self.spin_cm_h.value(),
            "crop_ratio": crop,
            "grid_cols": self.spin_cols.value(),
            "start_col": start_col,
            "start_row": self.spin_start_row.value(),
            "placement": "in_cell" if self.combo_placement.currentIndex() == 1 else "over",
        }

        self.btn_insert.setEnabled(False)
        self.progress.setValue(0)
        self.worker = InsertWorker(params)
        self.worker.progress.connect(self.progress.setValue)
        self.worker.status.connect(self.lbl_status.setText)
        self.worker.finished.connect(self._on_finished)
        self.worker.start()

    def _on_finished(self, error):
        self.btn_insert.setEnabled(True)
        if error:
            self.lbl_status.setText(f"Error: {error}")
            QMessageBox.critical(self, "Error", error)
        else:
            self.progress.setValue(100)
            self.lbl_status.setText("Done!")
            QMessageBox.information(self, "Success", "Images inserted successfully!")


def main():
    app = QApplication(sys.argv)
    app.setStyle("Fusion")
    win = MainWindow()
    win.show()
    sys.exit(app.exec_())


if __name__ == "__main__":
    main()
