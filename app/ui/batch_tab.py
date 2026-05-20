"""Batch Processor tab — folder-to-folder image transformations.

Lives next to the Excel Inserter tab inside MainWindow's QStackedWidget. Owns
its own widgets, builds a config dict from the current UI state, and hands
that dict to ``BatchProcessorWorker``.
"""

from __future__ import annotations

from datetime import datetime, time as dtime
from pathlib import Path

from PyQt5.QtCore import Qt, QDate, QTime
from PyQt5.QtGui import QColor, QPixmap, QIcon
from PyQt5.QtWidgets import (
    QWidget, QVBoxLayout, QHBoxLayout, QGridLayout,
    QButtonGroup, QFileDialog, QMessageBox, QStackedWidget, QColorDialog,
    QSizePolicy,
)

from qfluentwidgets import (
    PushButton, PrimaryPushButton,
    LineEdit, ComboBox, CheckBox, RadioButton, SpinBox, Slider, ProgressBar,
    BodyLabel, StrongBodyLabel, CaptionLabel,
    CardWidget, FluentIcon, DatePicker, TimePicker,
)

from app.core.models import CROP_PRESETS
from app.core.batch_processor import (
    BatchProcessorWorker, DATE_FORMATS, COLOR_PRESETS, POSITIONS,
)


# Local helper — matches the one in main_window.py. Kept inline to avoid
# coupling batch_tab to main_window's internals.
def _make_card(title: str = "") -> tuple:
    card = CardWidget()
    outer = QVBoxLayout(card)
    outer.setContentsMargins(10, 8, 10, 8)
    outer.setSpacing(6)
    if title:
        outer.addWidget(StrongBodyLabel(title))
    return card, outer


FONT_CHOICES = ["Arial", "Helvetica", "Courier", "Roboto", "Times New Roman"]


class BatchProcessorTab(QWidget):
    """Single Qt widget hosting the Batch Processor UI."""

    def __init__(self, parent=None):
        super().__init__(parent)
        self._worker: BatchProcessorWorker | None = None
        self._wm_color: tuple[int, int, int] = COLOR_PRESETS["White"]
        self._build_ui()

    # ── UI construction ───────────────────────────────────────────────────
    def _build_ui(self):
        root = QVBoxLayout(self)
        root.setSpacing(6)
        root.setContentsMargins(0, 0, 0, 0)

        self._build_io_card(root)
        self._build_resize_card(root)
        self._build_crop_card(root)
        self._build_metadata_card(root)
        self._build_watermark_card(root)
        self._build_action_row(root)
        root.addStretch(1)

        # Make sure conditional rows reflect initial selection
        self._sync_resize_visibility()
        self._sync_crop_custom_visibility()
        self._sync_watermark_mode_visibility()
        self._on_same_date_toggled(False)
        self._on_wm_same_as_metadata_toggled(False)
        self._on_overwrite_toggled(False)

    # ── Input / output ────────────────────────────────────────────────────
    def _build_io_card(self, root):
        card, lay = _make_card("Input / Output")

        in_row = QHBoxLayout()
        in_row.addWidget(BodyLabel("Input folder:"))
        self.le_input = LineEdit()
        self.le_input.setReadOnly(True)
        self.le_input.setPlaceholderText("Folder containing images")
        self.btn_browse_input = PushButton(FluentIcon.FOLDER, "Browse...")
        self.btn_browse_input.clicked.connect(self._browse_input)
        in_row.addWidget(self.le_input, 1)
        in_row.addWidget(self.btn_browse_input)
        lay.addLayout(in_row)

        out_row = QHBoxLayout()
        out_row.addWidget(BodyLabel("Output folder:"))
        self.le_output = LineEdit()
        self.le_output.setReadOnly(True)
        self.le_output.setPlaceholderText("Where processed images are written")
        self.btn_browse_output = PushButton(FluentIcon.FOLDER, "Browse...")
        self.btn_browse_output.clicked.connect(self._browse_output)
        out_row.addWidget(self.le_output, 1)
        out_row.addWidget(self.btn_browse_output)
        lay.addLayout(out_row)

        self.cb_overwrite = CheckBox("Overwrite originals (in-place)")
        self.cb_overwrite.toggled.connect(self._on_overwrite_toggled)
        lay.addWidget(self.cb_overwrite)

        self.lbl_overwrite_warn = CaptionLabel(
            "⚠ This permanently modifies the source files. There is no undo."
        )
        self.lbl_overwrite_warn.setStyleSheet("color: #e74c3c;")
        self.lbl_overwrite_warn.hide()
        lay.addWidget(self.lbl_overwrite_warn)

        root.addWidget(card)

    def _on_overwrite_toggled(self, checked):
        self.le_output.setEnabled(not checked)
        self.btn_browse_output.setEnabled(not checked)
        self.lbl_overwrite_warn.setVisible(checked)

    def _browse_input(self):
        path = QFileDialog.getExistingDirectory(self, "Select input folder")
        if path:
            self.le_input.setText(path)

    def _browse_output(self):
        path = QFileDialog.getExistingDirectory(self, "Select output folder")
        if path:
            self.le_output.setText(path)

    # ── Resize ────────────────────────────────────────────────────────────
    def _build_resize_card(self, root):
        card, lay = _make_card("Resize")
        bg = QButtonGroup(self)
        self.rb_resize_none = RadioButton("None")
        self.rb_resize_long = RadioButton("By long side (px)")
        self.rb_resize_pct = RadioButton("By percentage")
        self.rb_resize_exact = RadioButton("By exact dimensions (W×H)")
        self.rb_resize_none.setChecked(True)
        for rb in [self.rb_resize_none, self.rb_resize_long,
                   self.rb_resize_pct, self.rb_resize_exact]:
            bg.addButton(rb)
            rb.toggled.connect(self._sync_resize_visibility)

        row1 = QHBoxLayout()
        row1.addWidget(self.rb_resize_none)
        row1.addWidget(self.rb_resize_long)
        row1.addWidget(self.rb_resize_pct)
        row1.addWidget(self.rb_resize_exact)
        row1.addStretch()
        lay.addLayout(row1)

        # Conditional rows live in a stacked widget so we don't reflow.
        self.resize_stack = QStackedWidget()
        # Page 0: empty (None)
        self.resize_stack.addWidget(QWidget())
        # Page 1: long side
        long_page = QWidget(); long_lay = QHBoxLayout(long_page)
        long_lay.setContentsMargins(0, 0, 0, 0)
        self.spin_long_side = SpinBox()
        self.spin_long_side.setRange(100, 10000)
        self.spin_long_side.setValue(1920)
        long_lay.addWidget(BodyLabel("Long side:"))
        long_lay.addWidget(self.spin_long_side)
        long_lay.addWidget(BodyLabel("px"))
        long_lay.addStretch()
        self.resize_stack.addWidget(long_page)
        # Page 2: percent
        pct_page = QWidget(); pct_lay = QHBoxLayout(pct_page)
        pct_lay.setContentsMargins(0, 0, 0, 0)
        self.spin_resize_pct = SpinBox()
        self.spin_resize_pct.setRange(1, 500)
        self.spin_resize_pct.setValue(100)
        pct_lay.addWidget(BodyLabel("Scale:"))
        pct_lay.addWidget(self.spin_resize_pct)
        pct_lay.addWidget(BodyLabel("%"))
        pct_lay.addStretch()
        self.resize_stack.addWidget(pct_page)
        # Page 3: exact
        exact_page = QWidget(); exact_lay = QHBoxLayout(exact_page)
        exact_lay.setContentsMargins(0, 0, 0, 0)
        self.spin_resize_w = SpinBox()
        self.spin_resize_w.setRange(1, 20000)
        self.spin_resize_w.setValue(1920)
        self.spin_resize_h = SpinBox()
        self.spin_resize_h.setRange(1, 20000)
        self.spin_resize_h.setValue(1080)
        self.cb_keep_aspect = CheckBox("Keep aspect ratio")
        self.cb_keep_aspect.setChecked(True)
        exact_lay.addWidget(BodyLabel("W:"))
        exact_lay.addWidget(self.spin_resize_w)
        exact_lay.addWidget(BodyLabel("H:"))
        exact_lay.addWidget(self.spin_resize_h)
        exact_lay.addWidget(self.cb_keep_aspect)
        exact_lay.addStretch()
        self.resize_stack.addWidget(exact_page)

        lay.addWidget(self.resize_stack)
        root.addWidget(card)

    def _sync_resize_visibility(self, *_):
        if self.rb_resize_none.isChecked():
            self.resize_stack.setCurrentIndex(0)
        elif self.rb_resize_long.isChecked():
            self.resize_stack.setCurrentIndex(1)
        elif self.rb_resize_pct.isChecked():
            self.resize_stack.setCurrentIndex(2)
        else:
            self.resize_stack.setCurrentIndex(3)

    def _get_resize_mode(self) -> str:
        if self.rb_resize_long.isChecked():
            return "long_side"
        if self.rb_resize_pct.isChecked():
            return "percent"
        if self.rb_resize_exact.isChecked():
            return "exact"
        return "none"

    # ── Crop ──────────────────────────────────────────────────────────────
    def _build_crop_card(self, root):
        card, lay = _make_card("Crop")
        row = QHBoxLayout()
        row.addWidget(BodyLabel("Ratio:"))
        self.combo_crop = ComboBox()
        crop_items = list(CROP_PRESETS.keys()) + ["Custom W:H"]
        self.combo_crop.addItems(crop_items)
        self.combo_crop.currentTextChanged.connect(self._sync_crop_custom_visibility)
        row.addWidget(self.combo_crop)
        row.addStretch()
        lay.addLayout(row)

        self.crop_custom_row = QWidget()
        ccrow = QHBoxLayout(self.crop_custom_row)
        ccrow.setContentsMargins(0, 0, 0, 0)
        self.spin_crop_w = SpinBox()
        self.spin_crop_w.setRange(1, 100)
        self.spin_crop_w.setValue(5)
        self.spin_crop_h = SpinBox()
        self.spin_crop_h.setRange(1, 100)
        self.spin_crop_h.setValue(7)
        ccrow.addWidget(BodyLabel("W:"))
        ccrow.addWidget(self.spin_crop_w)
        ccrow.addWidget(BodyLabel("H:"))
        ccrow.addWidget(self.spin_crop_h)
        ccrow.addStretch()
        lay.addWidget(self.crop_custom_row)
        root.addWidget(card)

    def _sync_crop_custom_visibility(self, *_):
        self.crop_custom_row.setVisible(self.combo_crop.currentText() == "Custom W:H")

    def _get_crop_ratio(self):
        txt = self.combo_crop.currentText()
        if txt == "Custom W:H":
            return (self.spin_crop_w.value(), self.spin_crop_h.value())
        return CROP_PRESETS.get(txt)

    # ── Metadata ──────────────────────────────────────────────────────────
    def _build_metadata_card(self, root):
        card, lay = _make_card("Metadata")

        grid = QGridLayout()
        grid.setContentsMargins(0, 0, 0, 0)
        grid.setSpacing(6)

        grid.addWidget(BodyLabel("Created date:"), 0, 0)
        self.dp_created = DatePicker()
        self.tp_created = TimePicker()
        grid.addWidget(self.dp_created, 0, 1)
        grid.addWidget(self.tp_created, 0, 2)

        grid.addWidget(BodyLabel("Modified date:"), 1, 0)
        self.dp_modified = DatePicker()
        self.tp_modified = TimePicker()
        grid.addWidget(self.dp_modified, 1, 1)
        grid.addWidget(self.tp_modified, 1, 2)

        lay.addLayout(grid)

        self.cb_same_date = CheckBox("Same date for both")
        self.cb_same_date.toggled.connect(self._on_same_date_toggled)
        lay.addWidget(self.cb_same_date)

        self.cb_write_exif = CheckBox(
            "Also write to EXIF DateTimeOriginal (for photos)"
        )
        lay.addWidget(self.cb_write_exif)

        note = BodyLabel(
            "Note: Created date can only be modified on Windows. "
            "On macOS, only Modified date is changed."
        )
        note.setStyleSheet("color: gray; font-style: italic;")
        note.setWordWrap(True)
        lay.addWidget(note)

        root.addWidget(card)

    def _on_same_date_toggled(self, checked):
        # When checked, the modified picker mirrors the created picker and
        # is disabled.
        self.dp_modified.setEnabled(not checked)
        self.tp_modified.setEnabled(not checked)
        if checked:
            try:
                self.dp_modified.setDate(self.dp_created.getDate())
                self.tp_modified.setTime(self.tp_created.getTime())
            except Exception:
                pass

    def _read_datetime(self, dp: DatePicker, tp: TimePicker) -> datetime | None:
        try:
            qd = dp.getDate()
            qt = tp.getTime()
        except Exception:
            return None
        if qd is None or not qd.isValid():
            return None
        t = qt if (qt is not None and qt.isValid()) else QTime(0, 0)
        return datetime(qd.year(), qd.month(), qd.day(),
                        t.hour(), t.minute(), t.second())

    # ── Watermark ─────────────────────────────────────────────────────────
    def _build_watermark_card(self, root):
        card, lay = _make_card("Watermark")
        bg = QButtonGroup(self)
        self.rb_wm_none = RadioButton("None")
        self.rb_wm_date = RadioButton("Date")
        self.rb_wm_image = RadioButton("Image")
        self.rb_wm_none.setChecked(True)
        for rb in [self.rb_wm_none, self.rb_wm_date, self.rb_wm_image]:
            bg.addButton(rb)
            rb.toggled.connect(self._sync_watermark_mode_visibility)
        rb_row = QHBoxLayout()
        rb_row.addWidget(self.rb_wm_none)
        rb_row.addWidget(self.rb_wm_date)
        rb_row.addWidget(self.rb_wm_image)
        rb_row.addStretch()
        lay.addLayout(rb_row)

        self.wm_stack = QStackedWidget()
        # Page 0: empty
        self.wm_stack.addWidget(QWidget())
        # Page 1: date watermark
        self.wm_stack.addWidget(self._build_date_wm_page())
        # Page 2: image watermark
        self.wm_stack.addWidget(self._build_image_wm_page())
        lay.addWidget(self.wm_stack)

        root.addWidget(card)

    def _sync_watermark_mode_visibility(self, *_):
        if self.rb_wm_date.isChecked():
            self.wm_stack.setCurrentIndex(1)
        elif self.rb_wm_image.isChecked():
            self.wm_stack.setCurrentIndex(2)
        else:
            self.wm_stack.setCurrentIndex(0)

    def _build_date_wm_page(self) -> QWidget:
        page = QWidget()
        lay = QVBoxLayout(page)
        lay.setContentsMargins(0, 6, 0, 0)
        lay.setSpacing(6)

        # Format
        row_fmt = QHBoxLayout()
        row_fmt.addWidget(BodyLabel("Format:"))
        self.combo_date_format = ComboBox()
        self.combo_date_format.addItems([f[0] for f in DATE_FORMATS])
        self.combo_date_format.setMinimumWidth(280)
        row_fmt.addWidget(self.combo_date_format, 1)
        lay.addLayout(row_fmt)

        # Watermark date
        row_date = QHBoxLayout()
        row_date.addWidget(BodyLabel("Watermark date:"))
        self.dp_wm = DatePicker()
        self.tp_wm = TimePicker()
        row_date.addWidget(self.dp_wm)
        row_date.addWidget(self.tp_wm)
        row_date.addStretch()
        lay.addLayout(row_date)

        self.cb_wm_same_as_metadata = CheckBox("Same as metadata date")
        self.cb_wm_same_as_metadata.toggled.connect(self._on_wm_same_as_metadata_toggled)
        lay.addWidget(self.cb_wm_same_as_metadata)

        # Color
        color_row = QHBoxLayout()
        color_row.addWidget(BodyLabel("Color:"))
        self.combo_color = ComboBox()
        self.combo_color.addItems(list(COLOR_PRESETS.keys()) + ["Custom..."])
        self.combo_color.currentTextChanged.connect(self._on_color_changed)
        color_row.addWidget(self.combo_color)
        self.lbl_color_swatch = BodyLabel("")
        self.lbl_color_swatch.setFixedSize(28, 18)
        self._update_color_swatch()
        color_row.addWidget(self.lbl_color_swatch)
        color_row.addStretch()
        lay.addLayout(color_row)

        # Font
        font_row = QHBoxLayout()
        font_row.addWidget(BodyLabel("Font:"))
        self.combo_font = ComboBox()
        self.combo_font.addItems(FONT_CHOICES)
        font_row.addWidget(self.combo_font)
        font_row.addStretch()
        lay.addLayout(font_row)

        # Font size
        size_row = QHBoxLayout()
        size_row.addWidget(BodyLabel("Size (% of image width):"))
        self.slider_font_size = Slider(Qt.Horizontal)
        self.slider_font_size.setRange(1, 20)
        self.slider_font_size.setValue(5)
        self.slider_font_size.setMinimumWidth(180)
        self.lbl_font_size = CaptionLabel("5%")
        self.slider_font_size.valueChanged.connect(
            lambda v: self.lbl_font_size.setText(f"{v}%")
        )
        size_row.addWidget(self.slider_font_size, 1)
        size_row.addWidget(self.lbl_font_size)
        lay.addLayout(size_row)

        # Position
        pos_row = QHBoxLayout()
        pos_row.addWidget(BodyLabel("Position:"))
        self.combo_position = ComboBox()
        self.combo_position.addItems(POSITIONS)
        pos_row.addWidget(self.combo_position)
        pos_row.addStretch()
        lay.addLayout(pos_row)

        # Shadow
        self.cb_shadow = CheckBox("Add shadow/outline")
        lay.addWidget(self.cb_shadow)

        # Opacity
        op_row = QHBoxLayout()
        op_row.addWidget(BodyLabel("Opacity:"))
        self.slider_opacity = Slider(Qt.Horizontal)
        self.slider_opacity.setRange(0, 100)
        self.slider_opacity.setValue(100)
        self.slider_opacity.setMinimumWidth(180)
        self.lbl_opacity = CaptionLabel("100%")
        self.slider_opacity.valueChanged.connect(
            lambda v: self.lbl_opacity.setText(f"{v}%")
        )
        op_row.addWidget(self.slider_opacity, 1)
        op_row.addWidget(self.lbl_opacity)
        lay.addLayout(op_row)

        # Margin
        margin_row = QHBoxLayout()
        margin_row.addWidget(BodyLabel("Margin (px):"))
        self.spin_margin = SpinBox()
        self.spin_margin.setRange(0, 200)
        self.spin_margin.setValue(20)
        margin_row.addWidget(self.spin_margin)
        margin_row.addStretch()
        lay.addLayout(margin_row)

        return page

    def _build_image_wm_page(self) -> QWidget:
        page = QWidget()
        lay = QVBoxLayout(page)
        lay.setContentsMargins(0, 6, 0, 0)
        lay.setSpacing(6)

        path_row = QHBoxLayout()
        path_row.addWidget(BodyLabel("Watermark PNG:"))
        self.le_wm_image = LineEdit()
        self.le_wm_image.setPlaceholderText("Path to PNG (transparent recommended)")
        self.btn_browse_wm_image = PushButton(FluentIcon.PHOTO, "Browse...")
        self.btn_browse_wm_image.clicked.connect(self._browse_wm_image)
        path_row.addWidget(self.le_wm_image, 1)
        path_row.addWidget(self.btn_browse_wm_image)
        lay.addLayout(path_row)

        # Size
        size_row = QHBoxLayout()
        size_row.addWidget(BodyLabel("Size (% of image width):"))
        self.slider_wm_image_size = Slider(Qt.Horizontal)
        self.slider_wm_image_size.setRange(1, 50)
        self.slider_wm_image_size.setValue(15)
        self.slider_wm_image_size.setMinimumWidth(180)
        self.lbl_wm_image_size = CaptionLabel("15%")
        self.slider_wm_image_size.valueChanged.connect(
            lambda v: self.lbl_wm_image_size.setText(f"{v}%")
        )
        size_row.addWidget(self.slider_wm_image_size, 1)
        size_row.addWidget(self.lbl_wm_image_size)
        lay.addLayout(size_row)

        # Opacity (shared concept; separate slider so date/image stay independent)
        op_row = QHBoxLayout()
        op_row.addWidget(BodyLabel("Opacity:"))
        self.slider_wm_image_opacity = Slider(Qt.Horizontal)
        self.slider_wm_image_opacity.setRange(0, 100)
        self.slider_wm_image_opacity.setValue(100)
        self.slider_wm_image_opacity.setMinimumWidth(180)
        self.lbl_wm_image_opacity = CaptionLabel("100%")
        self.slider_wm_image_opacity.valueChanged.connect(
            lambda v: self.lbl_wm_image_opacity.setText(f"{v}%")
        )
        op_row.addWidget(self.slider_wm_image_opacity, 1)
        op_row.addWidget(self.lbl_wm_image_opacity)
        lay.addLayout(op_row)

        # Position
        pos_row = QHBoxLayout()
        pos_row.addWidget(BodyLabel("Position:"))
        self.combo_wm_image_position = ComboBox()
        self.combo_wm_image_position.addItems(POSITIONS)
        pos_row.addWidget(self.combo_wm_image_position)
        pos_row.addStretch()
        lay.addLayout(pos_row)

        # Margin
        margin_row = QHBoxLayout()
        margin_row.addWidget(BodyLabel("Margin (px):"))
        self.spin_wm_image_margin = SpinBox()
        self.spin_wm_image_margin.setRange(0, 200)
        self.spin_wm_image_margin.setValue(20)
        margin_row.addWidget(self.spin_wm_image_margin)
        margin_row.addStretch()
        lay.addLayout(margin_row)

        return page

    def _browse_wm_image(self):
        path, _ = QFileDialog.getOpenFileName(
            self, "Select watermark image", "",
            "PNG / images (*.png *.jpg *.jpeg *.bmp *.webp)"
        )
        if path:
            self.le_wm_image.setText(path)

    def _on_wm_same_as_metadata_toggled(self, checked):
        self.dp_wm.setEnabled(not checked)
        self.tp_wm.setEnabled(not checked)

    def _on_color_changed(self, text):
        if text == "Custom...":
            initial = QColor(*self._wm_color)
            picked = QColorDialog.getColor(initial, self, "Pick watermark color")
            if picked.isValid():
                self._wm_color = (picked.red(), picked.green(), picked.blue())
        elif text in COLOR_PRESETS:
            self._wm_color = COLOR_PRESETS[text]
        self._update_color_swatch()

    def _update_color_swatch(self):
        if not hasattr(self, "lbl_color_swatch"):
            return
        r, g, b = self._wm_color
        self.lbl_color_swatch.setStyleSheet(
            f"background-color: rgb({r},{g},{b}); border: 1px solid #888; border-radius: 3px;"
        )

    # ── Action row ────────────────────────────────────────────────────────
    def _build_action_row(self, root):
        self.progress = ProgressBar()
        self.progress.setValue(0)
        self.progress.setMaximumHeight(8)
        root.addWidget(self.progress)

        row = QHBoxLayout()
        self.lbl_status = CaptionLabel("Ready")
        row.addWidget(self.lbl_status, 1)
        self.btn_cancel = PushButton("Cancel")
        self.btn_cancel.hide()
        self.btn_cancel.clicked.connect(self._cancel_processing)
        row.addWidget(self.btn_cancel)
        self.btn_process = PrimaryPushButton(FluentIcon.ACCEPT, "Process Images")
        self.btn_process.setMinimumHeight(36)
        self.btn_process.clicked.connect(self._start_processing)
        row.addWidget(self.btn_process)
        root.addLayout(row)

    # ── Config builder ────────────────────────────────────────────────────
    def build_config(self) -> dict:
        """Build the config dict for BatchProcessorWorker from current UI state."""
        cfg: dict = {
            "input_dir": self.le_input.text().strip(),
            "output_dir": self.le_output.text().strip(),
            "overwrite": self.cb_overwrite.isChecked(),
            "resize_mode": self._get_resize_mode(),
            "resize_long_side": self.spin_long_side.value(),
            "resize_percent": self.spin_resize_pct.value(),
            "resize_w": self.spin_resize_w.value(),
            "resize_h": self.spin_resize_h.value(),
            "resize_keep_aspect": self.cb_keep_aspect.isChecked(),
            "crop_ratio": self._get_crop_ratio(),
        }

        created_dt = self._read_datetime(self.dp_created, self.tp_created)
        if self.cb_same_date.isChecked():
            modified_dt = created_dt
        else:
            modified_dt = self._read_datetime(self.dp_modified, self.tp_modified)
        cfg["created_dt"] = created_dt
        cfg["modified_dt"] = modified_dt
        cfg["write_exif"] = self.cb_write_exif.isChecked()

        # Watermark
        if self.rb_wm_date.isChecked():
            cfg["watermark_mode"] = "date"
            if self.cb_wm_same_as_metadata.isChecked():
                cfg["wm_date_dt"] = created_dt
            else:
                cfg["wm_date_dt"] = self._read_datetime(self.dp_wm, self.tp_wm)
            cfg["wm_date_format_index"] = self.combo_date_format.currentIndex()
            cfg["wm_color"] = self._wm_color
            cfg["wm_font"] = self.combo_font.currentText()
            cfg["wm_font_size_pct"] = self.slider_font_size.value()
            cfg["wm_position"] = self.combo_position.currentText()
            cfg["wm_shadow"] = self.cb_shadow.isChecked()
            cfg["wm_opacity"] = self.slider_opacity.value()
            cfg["wm_margin"] = self.spin_margin.value()
        elif self.rb_wm_image.isChecked():
            cfg["watermark_mode"] = "image"
            cfg["wm_image_path"] = self.le_wm_image.text().strip()
            cfg["wm_image_size_pct"] = self.slider_wm_image_size.value()
            cfg["wm_opacity"] = self.slider_wm_image_opacity.value()
            cfg["wm_position"] = self.combo_wm_image_position.currentText()
            cfg["wm_margin"] = self.spin_wm_image_margin.value()
        else:
            cfg["watermark_mode"] = "none"

        return cfg

    # ── Processing control ────────────────────────────────────────────────
    def _validate_config(self, cfg: dict) -> str | None:
        in_dir = cfg.get("input_dir") or ""
        if not in_dir or not Path(in_dir).is_dir():
            return "Please pick a valid input folder."
        if not cfg.get("overwrite"):
            out_dir = cfg.get("output_dir") or ""
            if not out_dir:
                return "Please pick an output folder, or enable 'Overwrite originals'."
            if Path(out_dir).resolve() == Path(in_dir).resolve():
                return ("Output folder is the same as input. "
                        "Enable 'Overwrite originals' if that's intended.")
        if cfg.get("watermark_mode") == "image":
            wmp = cfg.get("wm_image_path") or ""
            if not wmp or not Path(wmp).is_file():
                return "Please pick a watermark PNG file."
        return None

    def _start_processing(self):
        cfg = self.build_config()
        err = self._validate_config(cfg)
        if err:
            QMessageBox.warning(self, "Cannot start", err)
            return
        if cfg.get("overwrite"):
            res = QMessageBox.question(
                self, "Overwrite originals?",
                "Source files will be permanently modified. Continue?",
                QMessageBox.Yes | QMessageBox.No
            )
            if res != QMessageBox.Yes:
                return

        self.btn_process.setEnabled(False)
        self.btn_cancel.show()
        self.progress.setValue(0)
        self.lbl_status.setText("Starting...")

        self._worker = BatchProcessorWorker(cfg)
        self._worker.progress.connect(self._on_progress)
        self._worker.finished.connect(self._on_finished)
        self._worker.start()

    def _cancel_processing(self):
        if self._worker is not None:
            self._worker.cancel()
            self.lbl_status.setText("Cancelling...")

    def _on_progress(self, current, total, filename):
        if total > 0:
            self.progress.setValue(int(current / total * 100))
        self.lbl_status.setText(f"Processing {current}/{total}: {filename}")

    def _on_finished(self, processed, error_count, errors):
        self.btn_process.setEnabled(True)
        self.btn_cancel.hide()
        self.progress.setValue(100 if processed else 0)
        if error_count == 0:
            self.lbl_status.setText(f"Done. Processed {processed} images.")
            QMessageBox.information(
                self, "Batch complete",
                f"Successfully processed {processed} image(s)."
            )
        else:
            err_preview = "\n".join(f"• {n}: {m}" for n, m in errors[:10])
            more = f"\n…and {len(errors) - 10} more" if len(errors) > 10 else ""
            self.lbl_status.setText(
                f"Done with {error_count} error(s). Processed {processed}."
            )
            QMessageBox.warning(
                self, "Batch finished with errors",
                f"Processed {processed} image(s), {error_count} failed.\n\n"
                f"{err_preview}{more}"
            )
        self._worker = None
