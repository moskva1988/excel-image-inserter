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
    QFrame, QGridLayout, QSizePolicy, QStackedWidget,
)
from PyQt5.QtCore import Qt, QSize, QThread, pyqtSignal
from PyQt5.QtGui import QPixmap, QIcon, QImage, QPainter, QPen, QColor, QFont

from PIL import Image as PILImage
import openpyxl
from openpyxl.drawing.image import Image as XLImage
from openpyxl.utils import get_column_letter


# ── Constants ──────────────────────────────────────────────────────────────────
CM_TO_EMU = 360000  # 1 cm ≈ 360000 EMU
CM_TO_PX_96 = 96 / 2.54  # pixels per cm at 96 DPI
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


# ── Worker thread ──────────────────────────────────────────────────────────────
class InsertWorker(QThread):
    progress = pyqtSignal(int)
    finished = pyqtSignal(str)  # empty = success, else error message
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
        # Open / create workbook
        if p["excel_path"] and os.path.exists(p["excel_path"]):
            wb = openpyxl.load_workbook(p["excel_path"])
        else:
            wb = openpyxl.Workbook()

        # Select / create sheet
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

            img = PILImage.open(img_path)
            img = img.convert("RGB")

            # Crop
            if p["crop_ratio"]:
                img = self._crop_center(img, p["crop_ratio"])

            # Resize by pixels (reduce file size)
            if p["resize_px_w"] or p["resize_px_h"]:
                img = self._resize_px(img, p["resize_px_w"], p["resize_px_h"])

            # Save to buffer
            buf = BytesIO()
            img.save(buf, format="JPEG", quality=90)
            buf.seek(0)

            xl_img = XLImage(buf)

            # Size in cm → Excel dimensions
            w_cm = p["display_w_cm"]
            h_cm = p["display_h_cm"]
            xl_img.width = w_cm * CM_TO_PX_96
            xl_img.height = h_cm * CM_TO_PX_96

            # Grid position
            row_offset = i // cols
            col_offset = i % cols

            img_w_px = xl_img.width
            img_h_px = xl_img.height

            if p["placement"] == "in_cell":
                cell_col = start_col_idx + col_offset
                cell_row = start_row + row_offset
                ws.column_dimensions[get_column_letter(cell_col)].width = w_cm * 4.8
                ws.row_dimensions[cell_row].height = h_cm * 28.35
                cell_ref = f"{get_column_letter(cell_col)}{cell_row}"
                ws.add_image(xl_img, cell_ref)
            else:
                # "Over cells" — space images using dedicated columns/rows
                # Each image gets its own column (width = image width)
                # and its own row (height = image height)
                # Gap columns/rows separate them
                gap_col_width = 1.5  # ~0.4 cm gap
                gap_row_height = 8   # ~3 mm gap

                # Column index: start + col_offset * 2 (image col + gap col)
                img_col = start_col_idx + col_offset * 2
                # Row index: start + row_offset * 2 (image row + gap row)
                img_row = start_row + row_offset * 2

                # Set column width for image column (Excel width ≈ pixels / 7.5)
                ws.column_dimensions[get_column_letter(img_col)].width = img_w_px / 7.5
                # Set gap column width
                if col_offset < cols - 1:
                    gap_col = img_col + 1
                    ws.column_dimensions[get_column_letter(gap_col)].width = gap_col_width

                # Set row height for image row (Excel height in points ≈ pixels * 0.75)
                ws.row_dimensions[img_row].height = img_h_px * 0.75
                # Set gap row height
                gap_row = img_row + 1
                ws.row_dimensions[gap_row].height = gap_row_height

                cell_ref = f"{get_column_letter(img_col)}{img_row}"
                ws.add_image(xl_img, cell_ref)
            self.progress.emit(int((i + 1) / total * 100))

        save_path = p["save_path"] or p["excel_path"]
        self.status.emit(f"Saving {save_path}...")
        wb.save(save_path)

    @staticmethod
    def _crop_center(img, ratio):
        w, h = img.size
        target_w, target_h = ratio
        target_aspect = target_w / target_h
        current_aspect = w / h

        if current_aspect > target_aspect:
            new_w = int(h * target_aspect)
            left = (w - new_w) // 2
            img = img.crop((left, 0, left + new_w, h))
        else:
            new_h = int(w / target_aspect)
            top = (h - new_h) // 2
            img = img.crop((0, top, w, top + new_h))
        return img

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


# ── Preview widget ─────────────────────────────────────────────────────────────
class GridPreview(QWidget):
    """Schematic preview of how images will be laid out."""

    def __init__(self):
        super().__init__()
        self.cols = 2
        self.rows = 2
        self.count = 4
        self.crop_ratio = None
        self.setMinimumSize(200, 150)
        self.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Expanding)

    def update_params(self, cols, rows, count, crop_ratio):
        self.cols = max(1, cols)
        self.rows = max(1, rows)
        self.count = count
        self.crop_ratio = crop_ratio
        self.update()

    def paintEvent(self, event):
        painter = QPainter(self)
        painter.setRenderHint(QPainter.Antialiasing)

        w = self.width() - 20
        h = self.height() - 20
        ox, oy = 10, 10

        # Background
        painter.fillRect(self.rect(), QColor("#1a1a2e"))
        painter.setPen(QPen(QColor("#333"), 1))
        painter.drawRect(ox, oy, w, h)

        if self.count == 0:
            return

        cols = self.cols
        actual_rows = math.ceil(self.count / cols) if cols else 1
        if self.rows > 0:
            actual_rows = min(actual_rows, self.rows)

        cell_w = w / cols
        cell_h = h / actual_rows
        gap = 3

        # Determine image aspect for drawing
        if self.crop_ratio:
            aspect = self.crop_ratio[0] / self.crop_ratio[1]
        else:
            aspect = 4 / 3  # default

        drawn = 0
        for r in range(actual_rows):
            for c in range(cols):
                if drawn >= self.count:
                    break
                cx = ox + c * cell_w + gap
                cy = oy + r * cell_h + gap
                cw = cell_w - gap * 2
                ch = cell_h - gap * 2

                # Fit image rect inside cell
                img_aspect = aspect
                cell_aspect = cw / ch if ch else 1
                if img_aspect > cell_aspect:
                    iw = cw
                    ih = cw / img_aspect
                else:
                    ih = ch
                    iw = ch * img_aspect
                ix = cx + (cw - iw) / 2
                iy = cy + (ch - ih) / 2

                # Draw cell border
                painter.setPen(QPen(QColor("#444"), 1, Qt.DashLine))
                painter.drawRect(int(cx), int(cy), int(cw), int(ch))

                # Draw image placeholder
                color = QColor("#6366f1") if drawn < self.count else QColor("#333")
                painter.fillRect(int(ix), int(iy), int(iw), int(ih), color)
                painter.setPen(QPen(QColor("#fff"), 1))
                painter.setFont(QFont("Arial", 8))
                painter.drawText(int(ix), int(iy), int(iw), int(ih),
                                 Qt.AlignCenter, str(drawn + 1))
                drawn += 1

        painter.end()


# ── Main window ────────────────────────────────────────────────────────────────
class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Excel Image Inserter")
        self.setMinimumSize(750, 700)
        self.image_paths = []
        self._build_ui()

    def _build_ui(self):
        central = QWidget()
        self.setCentralWidget(central)
        root = QVBoxLayout(central)
        root.setSpacing(8)

        # ── Excel file ─────────────────────────────────────────────────────
        grp_file = QGroupBox("Excel File")
        lay_file = QVBoxLayout(grp_file)

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

        # Sheet selector
        row3 = QHBoxLayout()
        row3.addWidget(QLabel("Sheet:"))
        self.combo_sheet = QComboBox()
        self.combo_sheet.setMinimumWidth(150)
        row3.addWidget(self.combo_sheet, 1)
        self.cb_new_sheet = QCheckBox("New sheet:")
        self.le_new_sheet = QLineEdit()
        self.le_new_sheet.setPlaceholderText("Sheet name")
        self.le_new_sheet.setEnabled(False)
        self.cb_new_sheet.toggled.connect(self.le_new_sheet.setEnabled)
        self.cb_new_sheet.toggled.connect(lambda v: self.combo_sheet.setEnabled(not v))
        row3.addWidget(self.cb_new_sheet)
        row3.addWidget(self.le_new_sheet)
        lay_file.addLayout(row3)

        self.rb_new.toggled.connect(self._on_file_mode_changed)
        self.rb_open.toggled.connect(self._on_file_mode_changed)

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

        # View mode buttons (right side)
        self.btn_view_thumb = QPushButton("Thumbnails")
        self.btn_view_detail = QPushButton("Details")
        self.btn_view_grid = QPushButton("Grid")
        for b in [self.btn_view_thumb, self.btn_view_detail, self.btn_view_grid]:
            b.setCheckable(True)
            b.setMaximumWidth(80)
            b.setStyleSheet("QPushButton:checked{background:#6366f1;color:#fff;border-radius:4px}")
        self.btn_view_thumb.setChecked(True)
        self.btn_view_thumb.clicked.connect(lambda: self._switch_view("thumb"))
        self.btn_view_detail.clicked.connect(lambda: self._switch_view("detail"))
        self.btn_view_grid.clicked.connect(lambda: self._switch_view("grid"))
        btn_row.addWidget(self.btn_view_thumb)
        btn_row.addWidget(self.btn_view_detail)
        btn_row.addWidget(self.btn_view_grid)
        lay_img.addLayout(btn_row)

        # Thumbnail list view
        self.list_images = QListWidget()
        self.list_images.setIconSize(QSize(64, 64))
        self.list_images.setSelectionMode(QAbstractItemView.ExtendedSelection)
        self.list_images.setDragDropMode(QAbstractItemView.InternalMove)
        lay_img.addWidget(self.list_images)

        # Detail list view (with file sizes)
        self.list_details = QListWidget()
        self.list_details.setSelectionMode(QAbstractItemView.ExtendedSelection)
        self.list_details.setFont(QFont("Courier", 11))
        self.list_details.hide()
        lay_img.addWidget(self.list_details)

        # Grid preview view
        self.grid_preview = GridPreview()
        self.grid_preview.setMinimumHeight(180)
        self.grid_preview.hide()
        lay_img.addWidget(self.grid_preview)

        # Bottom row: count + total size
        bottom_row = QHBoxLayout()
        self.lbl_img_count = QLabel("0 images")
        self.lbl_total_size = QLabel("")
        self.lbl_total_size.setStyleSheet("color:#999")
        bottom_row.addWidget(self.lbl_img_count)
        bottom_row.addStretch()
        bottom_row.addWidget(self.lbl_total_size)
        lay_img.addLayout(bottom_row)

        root.addWidget(grp_img)

        # ── Settings (splitter: left=params, right=preview) ────────────────
        splitter = QSplitter(Qt.Horizontal)

        # Left: settings
        settings_w = QWidget()
        lay_settings = QVBoxLayout(settings_w)
        lay_settings.setContentsMargins(0, 0, 0, 0)

        # Resize pixels
        grp_resize = QGroupBox("Resize (pixels) — reduce file size")
        g_resize = QGridLayout(grp_resize)
        g_resize.addWidget(QLabel("Max width:"), 0, 0)
        self.spin_px_w = QSpinBox()
        self.spin_px_w.setRange(0, 10000)
        self.spin_px_w.setValue(1200)
        self.spin_px_w.setSpecialValueText("Auto")
        g_resize.addWidget(self.spin_px_w, 0, 1)
        g_resize.addWidget(QLabel("px"), 0, 2)

        g_resize.addWidget(QLabel("Max height:"), 1, 0)
        self.spin_px_h = QSpinBox()
        self.spin_px_h.setRange(0, 10000)
        self.spin_px_h.setValue(0)
        self.spin_px_h.setSpecialValueText("Auto")
        g_resize.addWidget(self.spin_px_h, 1, 1)
        g_resize.addWidget(QLabel("px"), 1, 2)
        lay_settings.addWidget(grp_resize)

        # Display size (cm)
        grp_display = QGroupBox("Display size in Excel (cm)")
        g_display = QGridLayout(grp_display)
        g_display.addWidget(QLabel("Width:"), 0, 0)
        self.spin_cm_w = QDoubleSpinBox()
        self.spin_cm_w.setRange(0.5, 50)
        self.spin_cm_w.setValue(6.0)
        self.spin_cm_w.setSingleStep(0.5)
        self.spin_cm_w.setSuffix(" cm")
        g_display.addWidget(self.spin_cm_w, 0, 1)

        g_display.addWidget(QLabel("Height:"), 1, 0)
        self.spin_cm_h = QDoubleSpinBox()
        self.spin_cm_h.setRange(0.5, 50)
        self.spin_cm_h.setValue(4.5)
        self.spin_cm_h.setSingleStep(0.5)
        self.spin_cm_h.setSuffix(" cm")
        g_display.addWidget(self.spin_cm_h, 1, 1)
        lay_settings.addWidget(grp_display)

        # Crop
        grp_crop = QGroupBox("Crop")
        g_crop = QHBoxLayout(grp_crop)
        g_crop.addWidget(QLabel("Aspect ratio:"))
        self.combo_crop = QComboBox()
        self.combo_crop.addItems(CROP_PRESETS.keys())
        self.combo_crop.currentTextChanged.connect(self._on_settings_changed)
        g_crop.addWidget(self.combo_crop, 1)
        lay_settings.addWidget(grp_crop)

        # Grid
        grp_grid = QGroupBox("Grid layout")
        g_grid = QGridLayout(grp_grid)
        g_grid.addWidget(QLabel("Columns:"), 0, 0)
        self.spin_cols = QSpinBox()
        self.spin_cols.setRange(1, 20)
        self.spin_cols.setValue(2)
        self.spin_cols.valueChanged.connect(self._on_settings_changed)
        g_grid.addWidget(self.spin_cols, 0, 1)

        g_grid.addWidget(QLabel("Max rows:"), 1, 0)
        self.spin_rows = QSpinBox()
        self.spin_rows.setRange(0, 1000)
        self.spin_rows.setValue(0)
        self.spin_rows.setSpecialValueText("Unlimited")
        self.spin_rows.valueChanged.connect(self._on_settings_changed)
        g_grid.addWidget(self.spin_rows, 1, 1)
        lay_settings.addWidget(grp_grid)

        # Position
        grp_pos = QGroupBox("Position in Excel")
        g_pos = QGridLayout(grp_pos)
        g_pos.addWidget(QLabel("Start cell:"), 0, 0)
        row_pos = QHBoxLayout()
        self.le_start_col = QLineEdit("A")
        self.le_start_col.setMaximumWidth(40)
        self.spin_start_row = QSpinBox()
        self.spin_start_row.setRange(1, 1048576)
        self.spin_start_row.setValue(1)
        row_pos.addWidget(self.le_start_col)
        row_pos.addWidget(self.spin_start_row)
        g_pos.addLayout(row_pos, 0, 1)

        g_pos.addWidget(QLabel("Placement:"), 1, 0)
        self.combo_placement = QComboBox()
        self.combo_placement.addItems(["Over cells (free)", "In cell (fit to cell)"])
        g_pos.addWidget(self.combo_placement, 1, 1)
        lay_settings.addWidget(grp_pos)

        lay_settings.addStretch()
        root.addWidget(settings_w)

        # ── Action ─────────────────────────────────────────────────────────
        sep = QFrame()
        sep.setFrameShape(QFrame.HLine)
        root.addWidget(sep)

        self.progress = QProgressBar()
        self.progress.setValue(0)
        root.addWidget(self.progress)

        self.lbl_status = QLabel("Ready")
        root.addWidget(self.lbl_status)

        self.btn_insert = QPushButton("Insert Images")
        self.btn_insert.setMinimumHeight(40)
        self.btn_insert.setStyleSheet("font-size:14px;font-weight:bold;")
        self.btn_insert.clicked.connect(self._do_insert)
        root.addWidget(self.btn_insert)

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
                self._load_sheets(path)

    def _load_sheets(self, path):
        try:
            wb = openpyxl.load_workbook(path, read_only=True)
            self.combo_sheet.clear()
            self.combo_sheet.addItems(wb.sheetnames)
            wb.close()
        except Exception as e:
            QMessageBox.warning(self, "Error", f"Cannot read file:\n{e}")

    def _add_images(self):
        paths, _ = QFileDialog.getOpenFileNames(
            self, "Select Images", "",
            "Images (*.jpg *.jpeg *.png *.bmp *.webp *.tiff);;All Files (*)"
        )
        for p in paths:
            if p not in self.image_paths:
                self.image_paths.append(p)
                # Thumbnail list item
                item = QListWidgetItem(Path(p).name)
                try:
                    px = QPixmap(p).scaled(64, 64, Qt.KeepAspectRatio, Qt.SmoothTransformation)
                    item.setIcon(QIcon(px))
                except Exception:
                    pass
                item.setData(Qt.UserRole, p)
                self.list_images.addItem(item)
        self._rebuild_details()
        self._update_count()

    def _clear_images(self):
        self.list_images.clear()
        self.list_details.clear()
        self.image_paths.clear()
        self._update_count()

    def _remove_selected(self):
        # Get selected from whichever list is visible
        active_list = self.list_images if self.list_images.isVisible() else self.list_details
        selected_rows = sorted([active_list.row(it) for it in active_list.selectedItems()], reverse=True)
        for row in selected_rows:
            if 0 <= row < len(self.image_paths):
                self.image_paths.pop(row)
            self.list_images.takeItem(row)
            self.list_details.takeItem(row)
        self._update_count()

    def _rebuild_details(self):
        """Rebuild the detail list with file sizes and estimated sizes."""
        self.list_details.clear()
        max_w = self.spin_px_w.value() or None
        max_h = self.spin_px_h.value() or None
        for p in self.image_paths:
            size_mb = os.path.getsize(p) / (1024 * 1024)
            try:
                img = PILImage.open(p)
                w, h = img.size
                # Estimate resized size
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
                    # Rough JPEG estimate: ~0.5 bytes/pixel at q90
                    est_mb = new_pixels * 0.5 / (1024 * 1024)
                else:
                    est_mb = size_mb
                dim_str = f"{w}x{h}"
                text = f"{Path(p).name:<30} {dim_str:>10}  {size_mb:>6.1f} MB -> {est_mb:>5.1f} MB"
            except Exception:
                text = f"{Path(p).name:<30} {'?':>10}  {size_mb:>6.1f} MB"
            item = QListWidgetItem(text)
            item.setData(Qt.UserRole, p)
            self.list_details.addItem(item)

    def _switch_view(self, mode):
        self.btn_view_thumb.setChecked(mode == "thumb")
        self.btn_view_detail.setChecked(mode == "detail")
        self.btn_view_grid.setChecked(mode == "grid")
        self.list_images.setVisible(mode == "thumb")
        self.list_details.setVisible(mode == "detail")
        self.grid_preview.setVisible(mode == "grid")
        if mode == "detail":
            self._rebuild_details()
        if mode == "grid":
            self._on_settings_changed()

    def _update_count(self):
        n = len(self.image_paths)
        self.lbl_img_count.setText(f"{n} image{'s' if n != 1 else ''}")
        # Total original size
        total_mb = sum(os.path.getsize(p) / (1024 * 1024) for p in self.image_paths if os.path.exists(p))
        if total_mb > 0:
            self.lbl_total_size.setText(f"Total: {total_mb:.1f} MB")
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

    def _get_ordered_paths(self):
        """Return image paths in current list order (respects drag-drop)."""
        paths = []
        for i in range(self.list_images.count()):
            item = self.list_images.item(i)
            paths.append(item.data(Qt.UserRole))
        return paths

    def _do_insert(self):
        # Validate
        file_path = self.le_file.text().strip()
        if self.rb_open.isChecked() and (not file_path or not os.path.exists(file_path)):
            QMessageBox.warning(self, "Error", "Please select an existing Excel file.")
            return
        if self.rb_new.isChecked() and not file_path:
            QMessageBox.warning(self, "Error", "Please specify a file path to save.")
            return

        if self.list_images.count() == 0:
            QMessageBox.warning(self, "Error", "No images selected.")
            return

        sheet_new = self.cb_new_sheet.isChecked()
        if sheet_new:
            sheet_name = self.le_new_sheet.text().strip()
            if not sheet_name:
                QMessageBox.warning(self, "Error", "Please enter a name for the new sheet.")
                return
        else:
            sheet_name = self.combo_sheet.currentText()
            if not sheet_name:
                QMessageBox.warning(self, "Error", "Please select a sheet.")
                return

        start_col = self.le_start_col.text().strip().upper()
        if not start_col or not start_col.isalpha():
            QMessageBox.warning(self, "Error", "Start column must be a letter (e.g. A, B, C).")
            return

        crop_key = self.combo_crop.currentText()
        crop = CROP_PRESETS.get(crop_key)

        params = {
            "excel_path": file_path if self.rb_open.isChecked() else None,
            "save_path": file_path,
            "sheet_new": sheet_new,
            "sheet_name": sheet_name,
            "images": self._get_ordered_paths(),
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
