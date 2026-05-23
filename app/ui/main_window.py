import os
from pathlib import Path

from PyQt5.QtWidgets import (
    QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
    QFileDialog, QAbstractItemView,
    QButtonGroup, QMessageBox,
    QGridLayout, QSizePolicy,
    QTreeWidget, QTreeWidgetItem, QHeaderView,
    QInputDialog, QMenu, QStackedWidget,
)
from PyQt5.QtCore import Qt, QSize, QSettings
from PyQt5.QtGui import QPixmap, QIcon, QColor, QFont

import openpyxl

from qfluentwidgets import (
    PushButton, PrimaryPushButton, TransparentToolButton,
    LineEdit, ComboBox, EditableComboBox,
    CheckBox, RadioButton, SpinBox, DoubleSpinBox, ProgressBar,
    BodyLabel, StrongBodyLabel, TitleLabel, CaptionLabel,
    CardWidget, FluentIcon, setTheme, Theme, isDarkTheme, themeColor,
    SegmentedWidget,
)

from app.core.models import APP_VERSION, BUILD_NUMBER, CROP_PRESETS, GROUP_ICON, GROUP_ICON_COLLAPSED
from app.core.image_processor import estimate_size
from app.core.excel_writer import InsertWorker
from app.ui.grid_preview import GridPreview
from app.ui.image_list import ThumbStackView
from app.ui.batch_tab import BatchProcessorTab


def _make_card(title: str = "") -> tuple:
    """Create a CardWidget with optional bold title label.

    Returns (card_widget, inner_layout). The inner layout is a QVBoxLayout
    embedded in the card; if `title` is non-empty, a StrongBodyLabel header
    is added first.
    """
    card = CardWidget()
    outer = QVBoxLayout(card)
    outer.setContentsMargins(10, 8, 10, 8)
    outer.setSpacing(6)
    if title:
        header = StrongBodyLabel(title)
        outer.addWidget(header)
    return card, outer


# ── Main window ────────────────────────────────────────────────────────────────
class MainWindow(QMainWindow):
    GROUP_ROLE = Qt.UserRole + 1  # stores group index
    PATH_ROLE = Qt.UserRole + 2  # stores image path
    TYPE_ROLE = Qt.UserRole + 3  # "group" or "image"

    def __init__(self):
        super().__init__()
        self.setWindowTitle("Excel Image Inserter")
        self.setMinimumSize(750, 750)
        self.groups = [{"title": "Group 1", "images": []}]
        self._collapsed_groups = set()  # indices of collapsed groups
        self._settings = QSettings("ExcelImageInserter", "ExcelImageInserter")
        self._build_ui()

    @property
    def image_paths(self):
        return [p for g in self.groups for p in g["images"]]

    def _build_ui(self):
        central = QWidget()
        self.setCentralWidget(central)
        outer = QVBoxLayout(central)
        outer.setSpacing(6)

        # ── Top header (shared across tabs): title + theme + about ─────────
        file_header = QHBoxLayout()
        self.lbl_app_title = TitleLabel("Excel Image Inserter")
        file_header.addWidget(self.lbl_app_title)
        file_header.addStretch()
        file_header.addWidget(CaptionLabel("Theme:"))
        self.combo_theme = ComboBox()
        self.combo_theme.addItems(["System", "Light", "Dark"])
        self.combo_theme.setMinimumWidth(110)
        saved_theme = self._settings.value("ui/theme", "System")
        if saved_theme in ("System", "Light", "Dark"):
            self.combo_theme.setCurrentText(saved_theme)
        self.combo_theme.currentTextChanged.connect(self._on_theme_changed)
        file_header.addWidget(self.combo_theme)
        self.btn_about = TransparentToolButton(FluentIcon.HELP)
        self.btn_about.setFixedSize(28, 28)
        self.btn_about.setToolTip("About")
        self.btn_about.clicked.connect(self._show_about)
        file_header.addWidget(self.btn_about)
        outer.addLayout(file_header)

        # ── Tab switcher (SegmentedWidget) ─────────────────────────────────
        # SegmentedWidget chosen over Pivot for a cleaner two-tab toggle and
        # an addItem signature (routeKey, text, onClick) that maps directly
        # to a QStackedWidget index.
        self.tab_switcher = SegmentedWidget()
        self.tab_switcher.addItem("excel", "Excel Inserter",
                                  lambda: self.tab_stack.setCurrentIndex(0))
        self.tab_switcher.addItem("batch", "Batch Processor",
                                  lambda: self.tab_stack.setCurrentIndex(1))
        outer.addWidget(self.tab_switcher)

        # ── Stacked pages: Excel (existing) + Batch (new) ──────────────────
        self.tab_stack = QStackedWidget()
        outer.addWidget(self.tab_stack, 1)

        excel_page = QWidget()
        root = QVBoxLayout(excel_page)
        root.setSpacing(6)
        root.setContentsMargins(0, 0, 0, 0)
        self.tab_stack.addWidget(excel_page)

        self.batch_tab = BatchProcessorTab()
        self.tab_stack.addWidget(self.batch_tab)
        self.tab_switcher.setCurrentItem("excel")
        self.tab_stack.setCurrentIndex(0)

        # The "Excel File" title used to live in the top header. Now that
        # the top header is shared and shows the app name, surface a small
        # section heading on this page instead.
        root.addWidget(StrongBodyLabel("Excel File"))

        grp_file, lay_file = _make_card()
        lay_file.setSpacing(6)

        self.lbl_format = CaptionLabel("⚠ Only .xlsx (Excel 2007+) is supported. Old .xls files must be re-saved as .xlsx first.")
        self._apply_format_warning_style()
        self.lbl_format.setWordWrap(True)
        lay_file.addWidget(self.lbl_format)

        row1 = QHBoxLayout()
        self.rb_new = RadioButton("Create new")
        self.rb_open = RadioButton("Open existing")
        self.rb_open.setChecked(True)
        bg = QButtonGroup(self)
        bg.addButton(self.rb_new)
        bg.addButton(self.rb_open)
        row1.addWidget(self.rb_open)
        row1.addWidget(self.rb_new)
        lay_file.addLayout(row1)

        row2 = QHBoxLayout()
        self.le_file = LineEdit()
        self.le_file.setPlaceholderText("Path to .xlsx file")
        self.btn_browse_file = PushButton(FluentIcon.FOLDER, "Browse...")
        self.btn_browse_file.clicked.connect(self._browse_file)
        row2.addWidget(self.le_file, 1)
        row2.addWidget(self.btn_browse_file)
        lay_file.addLayout(row2)

        row3 = QHBoxLayout()
        row3.addWidget(BodyLabel("Sheet:"))
        self.combo_sheet = ComboBox()
        self.combo_sheet.setMinimumWidth(120)
        row3.addWidget(self.combo_sheet, 1)
        self.cb_new_sheet = CheckBox("New:")
        self.le_new_sheet = LineEdit()
        self.le_new_sheet.setPlaceholderText("Sheet name")
        self.le_new_sheet.setEnabled(False)
        self.cb_new_sheet.toggled.connect(self.le_new_sheet.setEnabled)
        self.cb_new_sheet.toggled.connect(lambda v: self.combo_sheet.setEnabled(not v))
        self.cb_new_sheet.toggled.connect(self._on_new_sheet_toggled)
        row3.addWidget(self.cb_new_sheet)
        row3.addWidget(self.le_new_sheet)
        lay_file.addLayout(row3)

        # Insert after selector (for new sheets)
        row_insert = QHBoxLayout()
        self.lbl_insert_after = BodyLabel("Insert after:")
        self.combo_insert_after = ComboBox()
        self.combo_insert_after.addItem("(at the end)")
        row_insert.addWidget(self.lbl_insert_after)
        row_insert.addWidget(self.combo_insert_after, 1)
        self.lbl_insert_after.hide()
        self.combo_insert_after.hide()
        lay_file.addLayout(row_insert)

        # TOC checkbox
        self.cb_toc = CheckBox("Create / update Contents sheet with links")
        self.cb_toc.setChecked(True)
        self.cb_toc.hide()
        lay_file.addWidget(self.cb_toc)

        self.rb_new.toggled.connect(self._on_file_mode_changed)
        root.addWidget(grp_file)

        # ── Images ─────────────────────────────────────────────────────────
        grp_img, lay_img = _make_card("Images")

        # Mode toggle
        mode_row = QHBoxLayout()
        self.cb_use_groups = CheckBox("Use groups (headers + TOC)")
        self.cb_use_groups.toggled.connect(self._on_group_mode_toggled)
        mode_row.addWidget(self.cb_use_groups)
        mode_row.addStretch()
        lay_img.addLayout(mode_row)

        # Image/group controls
        btn_row = QHBoxLayout()
        self.btn_add_img = PushButton(FluentIcon.PHOTO, "Add images...")
        self.btn_add_img.clicked.connect(self._add_images)
        self.btn_add_group = PushButton(FluentIcon.ADD, "Group")
        self.btn_add_group.clicked.connect(self._add_group)
        self.btn_add_group.hide()
        self.btn_remove = PushButton(FluentIcon.DELETE, "Remove")
        self.btn_remove.clicked.connect(self._remove_selected)
        self.btn_clear_img = PushButton(FluentIcon.BROOM, "Clear all")
        self.btn_clear_img.clicked.connect(self._clear_images)
        self.btn_move_up = PushButton("▲")
        self.btn_move_up.setMaximumWidth(34)
        self.btn_move_up.setToolTip("Move up")
        self.btn_move_up.clicked.connect(lambda: self._move_selected(-1))
        self.btn_move_down = PushButton("▼")
        self.btn_move_down.setMaximumWidth(34)
        self.btn_move_down.setToolTip("Move down")
        self.btn_move_down.clicked.connect(lambda: self._move_selected(1))
        btn_row.addWidget(self.btn_add_img)
        btn_row.addWidget(self.btn_add_group)
        btn_row.addWidget(self.btn_remove)
        btn_row.addWidget(self.btn_clear_img)
        btn_row.addWidget(self.btn_move_up)
        btn_row.addWidget(self.btn_move_down)
        btn_row.addStretch()

        # View switcher — keep plain QPushButton: PrimaryPushButton would conflict
        # with the main "Insert Images" action; PushButton lacks `setCheckable`
        # styling in Fluent. Plain Qt push buttons with checkable=True give us a
        # consistent toggle behaviour across themes.
        from PyQt5.QtWidgets import QPushButton as _ToggleButton
        self.btn_view_list = _ToggleButton("List")
        self.btn_view_detail = _ToggleButton("Details")
        self.btn_view_stack = _ToggleButton("Stack")
        for b in [self.btn_view_list, self.btn_view_detail, self.btn_view_stack]:
            b.setCheckable(True)
            b.setMaximumWidth(70)
        self._apply_view_toggle_styles()
        self.btn_view_list.setChecked(True)
        self.btn_view_list.clicked.connect(lambda: self._switch_view("list"))
        self.btn_view_detail.clicked.connect(lambda: self._switch_view("detail"))
        self.btn_view_stack.clicked.connect(lambda: self._switch_view("stack"))
        btn_row.addWidget(self.btn_view_list)
        btn_row.addWidget(self.btn_view_detail)
        btn_row.addWidget(self.btn_view_stack)
        lay_img.addLayout(btn_row)

        # Active group selector
        group_sel_row = QHBoxLayout()
        self.lbl_active_group = BodyLabel("Add to group:")
        self.lbl_active_group.hide()
        self.combo_active_group = ComboBox()
        self.combo_active_group.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Fixed)
        self.combo_active_group.hide()
        group_sel_row.addWidget(self.lbl_active_group)
        group_sel_row.addWidget(self.combo_active_group)
        lay_img.addLayout(group_sel_row)

        # View: List (thumbnails + groups)
        self.tree = QTreeWidget()
        self.tree.setHeaderLabels(["", "Name", "Size", "After", ""])
        self.tree.setIconSize(QSize(48, 48))
        self.tree.setSelectionMode(QAbstractItemView.ExtendedSelection)
        self.tree.setRootIsDecorated(False)
        self.tree.setColumnWidth(0, 56)
        self.tree.setColumnWidth(1, 200)
        self.tree.setColumnWidth(2, 70)
        self.tree.setColumnWidth(3, 70)
        self.tree.setColumnWidth(4, 30)
        self.tree.header().setStretchLastSection(False)
        self.tree.header().setSectionResizeMode(1, QHeaderView.Stretch)
        self.tree.setMinimumHeight(200)
        self.tree.itemClicked.connect(self._on_tree_click)
        self.tree.setDragDropMode(QAbstractItemView.InternalMove)
        self.tree.setDefaultDropAction(Qt.MoveAction)
        self.tree.setContextMenuPolicy(Qt.CustomContextMenu)
        self.tree.customContextMenuRequested.connect(self._on_tree_context_menu)
        lay_img.addWidget(self.tree)

        # View: Details (no thumbnails)
        self.tree_detail = QTreeWidget()
        self.tree_detail.setHeaderLabels(["", "Name", "Dimensions", "Size", "After", ""])
        self.tree_detail.setSelectionMode(QAbstractItemView.ExtendedSelection)
        self.tree_detail.setRootIsDecorated(False)
        self.tree_detail.setColumnWidth(0, 30)
        self.tree_detail.setColumnWidth(1, 200)
        self.tree_detail.setColumnWidth(2, 90)
        self.tree_detail.setColumnWidth(3, 70)
        self.tree_detail.setColumnWidth(4, 70)
        self.tree_detail.setColumnWidth(5, 30)
        self.tree_detail.header().setStretchLastSection(False)
        self.tree_detail.header().setSectionResizeMode(1, QHeaderView.Stretch)
        self.tree_detail.setMinimumHeight(200)
        self.tree_detail.itemClicked.connect(self._on_tree_detail_click)
        self.tree_detail.setContextMenuPolicy(Qt.CustomContextMenu)
        self.tree_detail.customContextMenuRequested.connect(self._on_tree_context_menu_detail)
        self.tree_detail.hide()
        lay_img.addWidget(self.tree_detail)

        # View: Stack (thumbnail cards)
        self.thumb_stack = ThumbStackView()
        self.thumb_stack.delete_requested.connect(self._delete_by_path_flat)
        self.thumb_stack.order_changed.connect(self._on_stack_reorder)
        self.thumb_stack.setMinimumHeight(200)
        self.thumb_stack.hide()
        lay_img.addWidget(self.thumb_stack)

        root.addWidget(grp_img, 1)

        # Image stats bar (outside the Images group box)
        stats_row = QHBoxLayout()
        stats_row.setContentsMargins(4, 0, 4, 0)
        self.lbl_img_count = BodyLabel("0 images")
        self.lbl_total_size = CaptionLabel("")
        stats_row.addWidget(self.lbl_img_count)
        stats_row.addStretch()
        stats_row.addWidget(self.lbl_total_size)
        root.addLayout(stats_row)

        # ── Settings: Resize + Display in one row ──────────────────────────
        settings_row = QHBoxLayout()

        grp_resize, lay_resize = _make_card("Resize (px)")
        grp_resize_inner = QWidget()
        g_r = QGridLayout(grp_resize_inner)
        g_r.setContentsMargins(0, 0, 0, 0)
        g_r.setSpacing(4)
        resize_presets = ["Auto", "64", "128", "256", "320", "480", "640", "800", "1024", "1200", "1600", "1920", "2048", "3840"]
        g_r.addWidget(BodyLabel("W:"), 0, 0)
        self.combo_px_w = EditableComboBox()
        self.combo_px_w.addItems(resize_presets)
        self.combo_px_w.setCurrentText("1200")
        self.combo_px_w.currentTextChanged.connect(self._on_resize_changed)
        g_r.addWidget(self.combo_px_w, 0, 1)
        g_r.addWidget(BodyLabel("H:"), 1, 0)
        self.combo_px_h = EditableComboBox()
        self.combo_px_h.addItems(resize_presets)
        self.combo_px_h.setCurrentText("Auto")
        self.combo_px_h.currentTextChanged.connect(self._on_resize_changed)
        g_r.addWidget(self.combo_px_h, 1, 1)
        lay_resize.addWidget(grp_resize_inner)
        settings_row.addWidget(grp_resize)

        grp_display, lay_display = _make_card("Display (cm)")
        grp_display_inner = QWidget()
        g_d = QGridLayout(grp_display_inner)
        g_d.setContentsMargins(0, 0, 0, 0)
        g_d.setSpacing(4)
        g_d.addWidget(BodyLabel("Mode:"), 0, 0)
        self.combo_display_mode = ComboBox()
        self.combo_display_mode.addItems(["Per image", "Fixed ratio", "Manual"])
        self.combo_display_mode.setCurrentIndex(1)
        self.combo_display_mode.currentIndexChanged.connect(self._on_display_mode_changed)
        g_d.addWidget(self.combo_display_mode, 0, 1)
        g_d.addWidget(BodyLabel("W:"), 1, 0)
        self.spin_cm_w = DoubleSpinBox()
        self.spin_cm_w.setRange(0.5, 50)
        self.spin_cm_w.setValue(6.0)
        self.spin_cm_w.setSingleStep(0.5)
        self.spin_cm_w.setSuffix(" cm")
        self.spin_cm_w.valueChanged.connect(self._on_cm_w_changed)
        g_d.addWidget(self.spin_cm_w, 1, 1)
        g_d.addWidget(BodyLabel("H:"), 2, 0)
        self.spin_cm_h = DoubleSpinBox()
        self.spin_cm_h.setRange(0.5, 50)
        self.spin_cm_h.setValue(4.5)
        self.spin_cm_h.setSingleStep(0.5)
        self.spin_cm_h.setSuffix(" cm")
        self.spin_cm_h.valueChanged.connect(self._on_cm_h_changed)
        g_d.addWidget(self.spin_cm_h, 2, 1)

        # ── Anchor axis (W/H) toggle — visible in Per image / Fixed ratio ──
        self.lbl_anchor_axis = BodyLabel("Anchor:")
        g_d.addWidget(self.lbl_anchor_axis, 3, 0)
        anchor_row = QHBoxLayout()
        anchor_row.setContentsMargins(0, 0, 0, 0)
        anchor_row.setSpacing(6)
        self.rb_anchor_w = RadioButton("W")
        self.rb_anchor_h = RadioButton("H")
        self.rb_anchor_w.setChecked(True)
        self._anchor_axis_group = QButtonGroup(self)
        self._anchor_axis_group.addButton(self.rb_anchor_w)
        self._anchor_axis_group.addButton(self.rb_anchor_h)
        self.rb_anchor_w.toggled.connect(self._on_anchor_axis_changed)
        anchor_row.addWidget(self.rb_anchor_w)
        anchor_row.addWidget(self.rb_anchor_h)
        anchor_row.addStretch()
        g_d.addLayout(anchor_row, 3, 1)

        # ── Fixed-ratio aspect dropdown — visible only in Fixed ratio ──
        self.lbl_fixed_aspect = BodyLabel("Ratio:")
        g_d.addWidget(self.lbl_fixed_aspect, 4, 0)
        aspect_row = QHBoxLayout()
        aspect_row.setContentsMargins(0, 0, 0, 0)
        aspect_row.setSpacing(4)
        self.combo_fixed_aspect = ComboBox()
        self.combo_fixed_aspect.addItems(
            ["1:1", "4:3", "3:2", "16:9", "3:4", "2:3", "9:16", "Custom..."]
        )
        self.combo_fixed_aspect.setCurrentText("4:3")
        self.combo_fixed_aspect.currentTextChanged.connect(self._on_aspect_changed)
        self.spin_aspect_w = SpinBox()
        self.spin_aspect_w.setRange(1, 100)
        self.spin_aspect_w.setValue(4)
        self.spin_aspect_w.setMaximumWidth(60)
        self.spin_aspect_w.valueChanged.connect(self._on_aspect_changed)
        self.spin_aspect_h = SpinBox()
        self.spin_aspect_h.setRange(1, 100)
        self.spin_aspect_h.setValue(3)
        self.spin_aspect_h.setMaximumWidth(60)
        self.spin_aspect_h.valueChanged.connect(self._on_aspect_changed)
        aspect_row.addWidget(self.combo_fixed_aspect, 1)
        aspect_row.addWidget(self.spin_aspect_w)
        aspect_row.addWidget(BodyLabel(":"))
        aspect_row.addWidget(self.spin_aspect_h)
        # Custom spinboxes hidden until "Custom..." is selected
        self.spin_aspect_w.hide()
        self.spin_aspect_h.hide()
        g_d.addLayout(aspect_row, 4, 1)

        lay_display.addWidget(grp_display_inner)
        self._display_aspect = 6.0 / 4.5
        self._cm_updating = False
        settings_row.addWidget(grp_display)

        grp_crop, lay_crop = _make_card("Crop")
        self.combo_crop = ComboBox()
        self.combo_crop.addItems(CROP_PRESETS.keys())
        self.combo_crop.currentTextChanged.connect(self._on_settings_changed)
        self.combo_crop.currentTextChanged.connect(self._on_resize_changed)
        lay_crop.addWidget(self.combo_crop)
        lay_crop.addStretch()
        settings_row.addWidget(grp_crop)

        root.addLayout(settings_row)

        # ── Grid + Position + Preview ──────────────────────────────────────
        grid_row = QHBoxLayout()

        grp_grid, lay_grid = _make_card("Grid")
        grp_grid_inner = QWidget()
        g_g = QGridLayout(grp_grid_inner)
        g_g.setContentsMargins(0, 0, 0, 0)
        g_g.setSpacing(4)
        g_g.addWidget(BodyLabel("Cols:"), 0, 0)
        self.spin_cols = SpinBox()
        self.spin_cols.setRange(1, 20)
        self.spin_cols.setValue(2)
        self.spin_cols.valueChanged.connect(self._on_settings_changed)
        g_g.addWidget(self.spin_cols, 0, 1)
        g_g.addWidget(BodyLabel("H gap:"), 1, 0)
        self.spin_gap_h = DoubleSpinBox()
        self.spin_gap_h.setRange(0, 50)
        self.spin_gap_h.setValue(0.5)
        self.spin_gap_h.setSingleStep(0.05)
        self.spin_gap_h.setSuffix(" cm")
        self.spin_gap_h.setDecimals(2)
        g_g.addWidget(self.spin_gap_h, 1, 1)
        g_g.addWidget(BodyLabel("V gap:"), 2, 0)
        self.spin_gap_v = DoubleSpinBox()
        self.spin_gap_v.setRange(0, 50)
        self.spin_gap_v.setValue(0.5)
        self.spin_gap_v.setSingleStep(0.05)
        self.spin_gap_v.setSuffix(" cm")
        self.spin_gap_v.setDecimals(2)
        g_g.addWidget(self.spin_gap_v, 2, 1)
        lay_grid.addWidget(grp_grid_inner)
        grid_row.addWidget(grp_grid)

        grp_pos, lay_pos = _make_card("Position")
        grp_pos_inner = QWidget()
        g_p = QGridLayout(grp_pos_inner)
        g_p.setContentsMargins(0, 0, 0, 0)
        g_p.setSpacing(4)
        g_p.addWidget(BodyLabel("Cell:"), 0, 0)
        pos_row = QHBoxLayout()
        self.le_start_col = LineEdit()
        self.le_start_col.setText("A")
        self.le_start_col.setMaximumWidth(45)
        self.spin_start_row = SpinBox()
        self.spin_start_row.setRange(1, 1048576)
        self.spin_start_row.setValue(1)
        pos_row.addWidget(self.le_start_col)
        pos_row.addWidget(self.spin_start_row)
        g_p.addLayout(pos_row, 0, 1)
        g_p.addWidget(BodyLabel("Mode:"), 1, 0)
        self.combo_placement = ComboBox()
        self.combo_placement.addItems(["Over cells", "In cell"])
        self.combo_placement.currentIndexChanged.connect(self._on_settings_changed)
        self.le_start_col.textChanged.connect(self._on_settings_changed)
        self.spin_start_row.valueChanged.connect(self._on_settings_changed)
        g_p.addWidget(self.combo_placement, 1, 1)
        lay_pos.addWidget(grp_pos_inner)
        grid_row.addWidget(grp_pos)

        self.grid_preview = GridPreview()
        grid_row.addWidget(self.grid_preview, 1)

        root.addLayout(grid_row)

        # ── Action ─────────────────────────────────────────────────────────
        self.progress = ProgressBar()
        self.progress.setValue(0)
        self.progress.setMaximumHeight(8)
        root.addWidget(self.progress)

        action_row = QHBoxLayout()
        self.lbl_status = BodyLabel("Ready")
        action_row.addWidget(self.lbl_status, 1)
        self.btn_insert = PrimaryPushButton(FluentIcon.ACCEPT, "Insert Images")
        self.btn_insert.setMinimumHeight(36)
        self.btn_insert.clicked.connect(self._do_insert)
        action_row.addWidget(self.btn_insert)
        root.addLayout(action_row)

        # (About is shown via ? button in top header)

        # Apply initial visibility/enable state for the (new) default display
        # mode. setCurrentIndex above was called BEFORE the widgets it
        # toggles existed, so the signal-driven path can't be relied on here.
        self._on_display_mode_changed(self.combo_display_mode.currentIndex())

        self._rebuild_tree()

    # ── File/sheet management ─────────────────────────────────────────────
    def _on_file_mode_changed(self):
        is_open = self.rb_open.isChecked()
        self.btn_browse_file.setText("Browse..." if is_open else "Save as...")
        if not is_open:
            self.combo_sheet.clear()
            self.cb_new_sheet.setChecked(True)

    def _on_new_sheet_toggled(self, checked):
        self.lbl_insert_after.setVisible(checked and self.combo_sheet.count() > 0)
        self.combo_insert_after.setVisible(checked and self.combo_sheet.count() > 0)
        if checked:
            self.combo_insert_after.clear()
            self.combo_insert_after.addItem("(at the end)")
            for i in range(self.combo_sheet.count()):
                self.combo_insert_after.addItem(self.combo_sheet.itemText(i))

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

    # ── Group mode toggle ─────────────────────────────────────────────────
    def _on_group_mode_toggled(self, enabled):
        self.btn_add_group.setVisible(enabled)
        self.cb_toc.setVisible(enabled)
        self.lbl_active_group.setVisible(enabled)
        self.combo_active_group.setVisible(enabled)
        if not enabled:
            all_images = self.image_paths
            self.groups = [{"title": "All Images", "images": all_images}]
        self._rebuild_tree()

    # ── Group management ──────────────────────────────────────────────────
    def _add_group(self):
        name, ok = QInputDialog.getText(self, "New Group", "Group title:")
        if ok and name.strip():
            self.groups.append({"title": name.strip(), "images": []})
            self._rebuild_tree()
            # Auto-select newly created group
            self.combo_active_group.setCurrentIndex(len(self.groups) - 1)
            self._on_settings_changed()

    def _get_selected_group_idx(self):
        """Get group index of currently selected item."""
        items = self.tree.selectedItems()
        if not items:
            return len(self.groups) - 1 if self.groups else -1
        item = items[0]
        tp = item.data(0, self.TYPE_ROLE)
        if tp == "group":
            return item.data(0, self.GROUP_ROLE)
        elif tp == "image":
            return item.data(0, self.GROUP_ROLE)
        return 0

    # ── Tree view ─────────────────────────────────────────────────────────
    def _rebuild_tree(self):
        max_w, max_h = self._get_resize_px()
        use_groups = self.cb_use_groups.isChecked()

        # ── Update group selector combo ──
        prev_idx = self.combo_active_group.currentIndex()
        self.combo_active_group.blockSignals(True)
        self.combo_active_group.clear()
        for gi, group in enumerate(self.groups):
            self.combo_active_group.addItem(group["title"], userData=gi)
        if 0 <= prev_idx < len(self.groups):
            self.combo_active_group.setCurrentIndex(prev_idx)
        elif self.groups:
            self.combo_active_group.setCurrentIndex(len(self.groups) - 1)
        self.combo_active_group.blockSignals(False)

        # ── List view ──
        self.tree.clear()
        for gi, group in enumerate(self.groups):
            if use_groups:
                collapsed = gi in self._collapsed_groups
                icon = GROUP_ICON_COLLAPSED if collapsed else GROUP_ICON
                grp_item = QTreeWidgetItem([
                    icon,
                    f"{group['title']} ({len(group['images'])})",
                    "", "", ""
                ])
                grp_item.setData(0, self.TYPE_ROLE, "group")
                grp_item.setData(0, self.GROUP_ROLE, gi)
                grp_font = grp_item.font(1)
                grp_font.setBold(True)
                grp_item.setFont(1, grp_font)
                base = self.palette().color(self.backgroundRole())
                is_dark = base.lightnessF() < 0.5
                if is_dark:
                    grp_bg = QColor(base.red() + (255 - base.red()) // 5,
                                    base.green() + (255 - base.green()) // 5,
                                    base.blue() + (255 - base.blue()) // 5)
                else:
                    grp_bg = QColor(base.red() - base.red() // 10,
                                    base.green() - base.green() // 10,
                                    base.blue() - base.blue() // 10)
                grp_fg = QColor(Qt.white) if is_dark else QColor(Qt.black)
                for c in range(5):
                    grp_item.setBackground(c, grp_bg)
                    grp_item.setForeground(c, grp_fg)
                self.tree.addTopLevelItem(grp_item)
                if collapsed:
                    continue

            for p in group["images"]:
                orig_mb, est_mb, w, h = estimate_size(p, max_w, max_h)
                item = QTreeWidgetItem(["", Path(p).name, f"{orig_mb:.2f} MB", f"{est_mb:.2f} MB", "×"])
                item.setData(0, self.TYPE_ROLE, "image")
                item.setData(0, self.PATH_ROLE, p)
                item.setData(0, self.GROUP_ROLE, gi)
                try:
                    px = QPixmap(p).scaled(48, 48, Qt.KeepAspectRatio, Qt.SmoothTransformation)
                    item.setIcon(0, QIcon(px))
                except Exception:
                    pass
                self.tree.addTopLevelItem(item)

        # ── Detail view ──
        self.tree_detail.clear()
        for gi, group in enumerate(self.groups):
            if use_groups:
                collapsed = gi in self._collapsed_groups
                icon = GROUP_ICON_COLLAPSED if collapsed else GROUP_ICON
                grp_item = QTreeWidgetItem([icon, f"{group['title']} ({len(group['images'])})", "", "", "", ""])
                grp_item.setData(0, self.TYPE_ROLE, "group")
                grp_item.setData(0, self.GROUP_ROLE, gi)
                grp_font = grp_item.font(1)
                grp_font.setBold(True)
                grp_item.setFont(1, grp_font)
                base = self.palette().color(self.backgroundRole())
                is_dark = base.lightnessF() < 0.5
                if is_dark:
                    grp_bg = QColor(base.red() + (255 - base.red()) // 5,
                                    base.green() + (255 - base.green()) // 5,
                                    base.blue() + (255 - base.blue()) // 5)
                else:
                    grp_bg = QColor(base.red() - base.red() // 10,
                                    base.green() - base.green() // 10,
                                    base.blue() - base.blue() // 10)
                grp_fg = QColor(Qt.white) if is_dark else QColor(Qt.black)
                for c in range(6):
                    grp_item.setBackground(c, grp_bg)
                    grp_item.setForeground(c, grp_fg)
                self.tree_detail.addTopLevelItem(grp_item)
                if collapsed:
                    continue
            for p in group["images"]:
                orig_mb, est_mb, w, h = estimate_size(p, max_w, max_h)
                dim = f"{w}×{h}" if w else "?"
                item = QTreeWidgetItem(["", Path(p).name, dim, f"{orig_mb:.2f} MB", f"{est_mb:.2f} MB", "×"])
                item.setData(0, self.TYPE_ROLE, "image")
                item.setData(0, self.PATH_ROLE, p)
                item.setData(0, self.GROUP_ROLE, gi)
                self.tree_detail.addTopLevelItem(item)

        # ── Stack view ──
        all_images = self.image_paths
        self.thumb_stack.set_images(all_images, max_w, max_h)

        self._update_count()

    def _switch_view(self, mode):
        self.btn_view_list.setChecked(mode == "list")
        self.btn_view_detail.setChecked(mode == "detail")
        self.btn_view_stack.setChecked(mode == "stack")
        self.tree.setVisible(mode == "list")
        self.tree_detail.setVisible(mode == "detail")
        self.thumb_stack.setVisible(mode == "stack")

    def _on_tree_detail_click(self, item, col):
        tp = item.data(0, self.TYPE_ROLE)
        if tp == "group":
            gi = item.data(0, self.GROUP_ROLE)
            if col <= 1:
                if gi in self._collapsed_groups:
                    self._collapsed_groups.discard(gi)
                else:
                    self._collapsed_groups.add(gi)
                self._rebuild_tree()
        elif tp == "image" and col == 5:
            path = item.data(0, self.PATH_ROLE)
            if path:
                self._delete_by_path(path, item.data(0, self.GROUP_ROLE))

    def _on_tree_context_menu_detail(self, pos):
        item = self.tree_detail.itemAt(pos)
        if not item:
            return
        tp = item.data(0, self.TYPE_ROLE)
        menu = QMenu(self)
        if tp == "group":
            gi = item.data(0, self.GROUP_ROLE)
            menu.addAction("Rename group", lambda: self._rename_group(gi))
            if len(self.groups) > 1:
                menu.addAction("Delete group", lambda: self._delete_group(gi))
        elif tp == "image":
            path = item.data(0, self.PATH_ROLE)
            gi = item.data(0, self.GROUP_ROLE)
            if self.cb_use_groups.isChecked() and len(self.groups) > 1:
                move_menu = menu.addMenu("Move to group...")
                for i, g in enumerate(self.groups):
                    if i != gi:
                        move_menu.addAction(g["title"], lambda p=path, src=gi, dst=i: self._move_image_to_group(p, src, dst))
            menu.addAction("Remove", lambda: self._delete_by_path(path, gi))
        menu.exec_(self.tree_detail.viewport().mapToGlobal(pos))

    def _delete_by_path_flat(self, path):
        """Delete from stack view — find which group has it."""
        for gi, g in enumerate(self.groups):
            if path in g["images"]:
                self._delete_by_path(path, gi)
                return

    def _on_stack_reorder(self, new_order):
        """Reorder from stack view — applies to first group only in flat mode."""
        if not self.cb_use_groups.isChecked() and len(self.groups) == 1:
            self.groups[0]["images"] = new_order
            self._rebuild_tree()

    def _on_tree_click(self, item, col):
        tp = item.data(0, self.TYPE_ROLE)
        if tp == "group":
            gi = item.data(0, self.GROUP_ROLE)
            if col <= 1:
                # Toggle expand/collapse
                if gi in self._collapsed_groups:
                    self._collapsed_groups.discard(gi)
                else:
                    self._collapsed_groups.add(gi)
                self._rebuild_tree()
            return
        if tp == "image" and col == 4:
            path = item.data(0, self.PATH_ROLE)
            if path:
                self._delete_by_path(path, item.data(0, self.GROUP_ROLE))

    def _on_tree_context_menu(self, pos):
        item = self.tree.itemAt(pos)
        if not item:
            return
        tp = item.data(0, self.TYPE_ROLE)
        menu = QMenu(self)

        if tp == "group":
            gi = item.data(0, self.GROUP_ROLE)
            menu.addAction("Rename group", lambda: self._rename_group(gi))
            if len(self.groups) > 1:
                menu.addAction("Delete group", lambda: self._delete_group(gi))
            if gi > 0:
                menu.addAction("Move group up", lambda: self._move_group(gi, -1))
            if gi < len(self.groups) - 1:
                menu.addAction("Move group down", lambda: self._move_group(gi, 1))
        elif tp == "image":
            path = item.data(0, self.PATH_ROLE)
            gi = item.data(0, self.GROUP_ROLE)
            if self.cb_use_groups.isChecked() and len(self.groups) > 1:
                move_menu = menu.addMenu("Move to group...")
                for i, g in enumerate(self.groups):
                    if i != gi:
                        move_menu.addAction(g["title"], lambda p=path, src=gi, dst=i: self._move_image_to_group(p, src, dst))
            menu.addAction("Remove", lambda: self._delete_by_path(path, gi))

        menu.exec_(self.tree.viewport().mapToGlobal(pos))

    def _rename_group(self, gi):
        g = self.groups[gi]
        name, ok = QInputDialog.getText(self, "Rename Group", "New title:", text=g["title"])
        if ok and name.strip():
            g["title"] = name.strip()
            self._rebuild_tree()
            self._on_settings_changed()

    def _delete_group(self, gi):
        if len(self.groups) <= 1:
            return
        g = self.groups[gi]
        if g["images"]:
            if QMessageBox.question(self, "Delete Group",
                                    f"Delete '{g['title']}' with {len(g['images'])} images?",
                                    QMessageBox.Yes | QMessageBox.No) != QMessageBox.Yes:
                return
        self.groups.pop(gi)
        self._collapsed_groups.discard(gi)
        # Re-index collapsed groups
        self._collapsed_groups = {i - 1 if i > gi else i for i in self._collapsed_groups if i != gi}
        self._rebuild_tree()
        self._on_settings_changed()

    def _move_group(self, gi, direction):
        new_gi = gi + direction
        if new_gi < 0 or new_gi >= len(self.groups):
            return
        self.groups[gi], self.groups[new_gi] = self.groups[new_gi], self.groups[gi]
        # Update collapsed set
        new_collapsed = set()
        for c in self._collapsed_groups:
            if c == gi:
                new_collapsed.add(new_gi)
            elif c == new_gi:
                new_collapsed.add(gi)
            else:
                new_collapsed.add(c)
        self._collapsed_groups = new_collapsed
        self._rebuild_tree()
        self._on_settings_changed()

    def _move_image_to_group(self, path, src_gi, dst_gi):
        if path in self.groups[src_gi]["images"]:
            self.groups[src_gi]["images"].remove(path)
            self.groups[dst_gi]["images"].append(path)
            self._rebuild_tree()

    # ── Image management ──────────────────────────────────────────────────
    def _add_images(self):
        if self.cb_use_groups.isChecked():
            gi = self.combo_active_group.currentData()
            if gi is None:
                gi = 0
        else:
            gi = 0
        paths, _ = QFileDialog.getOpenFileNames(
            self, "Select Images", "",
            "Images (*.jpg *.jpeg *.png *.bmp *.webp *.tiff);;All Files (*)"
        )
        new_paths = [p for p in paths if p not in self.groups[gi]["images"]]
        if not new_paths:
            return
        self.groups[gi]["images"].extend(new_paths)
        # Make sure this group is expanded
        self._collapsed_groups.discard(gi)
        self._rebuild_tree()

    def _clear_images(self):
        total = sum(len(g["images"]) for g in self.groups)
        if total == 0:
            return
        if QMessageBox.question(self, "Clear", f"Remove all {total} images?",
                                QMessageBox.Yes | QMessageBox.No) != QMessageBox.Yes:
            return
        for g in self.groups:
            g["images"].clear()
        self._rebuild_tree()

    def _remove_selected(self):
        items = self.tree.selectedItems()
        if not items:
            return
        to_remove = []  # (gi, path) pairs
        groups_to_remove = []
        for item in items:
            tp = item.data(0, self.TYPE_ROLE)
            if tp == "image":
                to_remove.append((item.data(0, self.GROUP_ROLE), item.data(0, self.PATH_ROLE)))
            elif tp == "group" and self.cb_use_groups.isChecked():
                groups_to_remove.append(item.data(0, self.GROUP_ROLE))

        if not to_remove and not groups_to_remove:
            return

        desc = f"{len(to_remove)} image(s)" if to_remove else ""
        if groups_to_remove:
            desc += f"{', ' if desc else ''}{len(groups_to_remove)} group(s)"
        if QMessageBox.question(self, "Remove", f"Remove {desc}?",
                                QMessageBox.Yes | QMessageBox.No) != QMessageBox.Yes:
            return

        for gi, path in to_remove:
            if gi < len(self.groups) and path in self.groups[gi]["images"]:
                self.groups[gi]["images"].remove(path)

        for gi in sorted(groups_to_remove, reverse=True):
            if len(self.groups) > 1:
                self.groups.pop(gi)

        self._rebuild_tree()

    def _delete_by_path(self, path, gi):
        if QMessageBox.question(self, "Remove", f"Remove {Path(path).name}?",
                                QMessageBox.Yes | QMessageBox.No) != QMessageBox.Yes:
            return
        if gi < len(self.groups) and path in self.groups[gi]["images"]:
            self.groups[gi]["images"].remove(path)
        self._rebuild_tree()

    def _move_selected(self, direction):
        items = self.tree.selectedItems()
        if not items:
            return
        item = items[0]
        tp = item.data(0, self.TYPE_ROLE)

        if tp == "group" and self.cb_use_groups.isChecked():
            gi = item.data(0, self.GROUP_ROLE)
            self._move_group(gi, direction)
            return

        if tp == "image":
            gi = item.data(0, self.GROUP_ROLE)
            path = item.data(0, self.PATH_ROLE)
            if gi >= len(self.groups):
                return
            images = self.groups[gi]["images"]
            idx = images.index(path) if path in images else -1
            if idx < 0:
                return
            new_idx = idx + direction
            if new_idx < 0 or new_idx >= len(images):
                return
            images[idx], images[new_idx] = images[new_idx], images[idx]
            # Fast swap in tree
            tree_idx = self.tree.indexOfTopLevelItem(item)
            swap_idx = tree_idx + direction
            if 0 <= swap_idx < self.tree.topLevelItemCount():
                swap_item = self.tree.topLevelItem(swap_idx)
                if swap_item.data(0, self.TYPE_ROLE) == "image" and swap_item.data(0, self.GROUP_ROLE) == gi:
                    self.tree.blockSignals(True)
                    a = self.tree.takeTopLevelItem(max(tree_idx, swap_idx))
                    b = self.tree.takeTopLevelItem(min(tree_idx, swap_idx))
                    self.tree.insertTopLevelItem(min(tree_idx, swap_idx), a)
                    self.tree.insertTopLevelItem(max(tree_idx, swap_idx), b)
                    for i in range(self.tree.topLevelItemCount()):
                        if self.tree.topLevelItem(i).data(0, self.PATH_ROLE) == path:
                            self.tree.setCurrentItem(self.tree.topLevelItem(i))
                            break
                    self.tree.blockSignals(False)
                    return
            self._rebuild_tree()

    # ── Counts and settings ───────────────────────────────────────────────
    def _update_count(self):
        all_images = self.image_paths
        n = len(all_images)
        ng = len(self.groups)
        if self.cb_use_groups.isChecked():
            self.lbl_img_count.setText(f"{n} images in {ng} groups")
        else:
            self.lbl_img_count.setText(f"{n} image{'s' if n != 1 else ''}")
        max_w, max_h = self._get_resize_px()
        total_orig = sum(os.path.getsize(p) / (1024 * 1024) for p in all_images if os.path.exists(p))
        total_est = sum(estimate_size(p, max_w, max_h)[1] for p in all_images)
        if total_orig > 0:
            self.lbl_total_size.setText(f"Total: {total_orig:.2f} MB → {total_est:.2f} MB")
        else:
            self.lbl_total_size.setText("")
        self._on_settings_changed()

    def _get_resize_px(self):
        def _parse(combo):
            txt = combo.currentText().strip()
            if not txt or txt.lower() == "auto":
                return None
            try:
                return int(txt)
            except ValueError:
                return None
        return _parse(self.combo_px_w), _parse(self.combo_px_h)

    def _on_resize_changed(self, *_):
        if self.image_paths:
            self._rebuild_tree()

    def _apply_view_toggle_styles(self):
        """Apply theme-aware stylesheet to the List/Details/Stack toggle buttons.
        Recomputes the accent color so a Light↔Dark switch picks up the new tint."""
        try:
            accent = themeColor().name()
        except Exception:
            accent = "#6366f1"
        qss = (
            "QPushButton {"
            "  padding: 4px 8px;"
            "  border: 1px solid palette(mid);"
            "  border-radius: 4px;"
            "  background: transparent;"
            "  color: palette(text);"
            "}"
            "QPushButton:hover {"
            "  background: palette(midlight);"
            "}"
            f"QPushButton:checked {{"
            f"  background: {accent};"
            f"  color: white;"
            f"  border-color: {accent};"
            f"}}"
        )
        for b in (
            getattr(self, "btn_view_list", None),
            getattr(self, "btn_view_detail", None),
            getattr(self, "btn_view_stack", None),
        ):
            if b is not None:
                b.setStyleSheet(qss)

    def _apply_format_warning_style(self):
        """Theme-aware amber warning for the .xlsx-only notice."""
        if not hasattr(self, "lbl_format"):
            return
        # Amber works on both themes; brighten slightly on dark.
        color = "#f39c12" if isDarkTheme() else "#e67e22"
        self.lbl_format.setStyleSheet(f"color: {color}; padding: 2px 0;")

    def _refresh_themed_widgets(self):
        """Re-run all per-widget stylesheets that depend on the active theme.
        Called from _on_theme_changed so accent and warning colors flip."""
        self._apply_view_toggle_styles()
        self._apply_format_warning_style()

    def _on_theme_changed(self, value):
        """Apply and persist a Light/Dark/System theme choice."""
        mapping = {"Light": Theme.LIGHT, "Dark": Theme.DARK, "System": Theme.AUTO}
        setTheme(mapping.get(value, Theme.AUTO))
        self._settings.setValue("ui/theme", value)

        # Force Fluent to flush the application stylesheet to all open windows.
        # Re-applying the existing stylesheet triggers a global polish/repolish.
        try:
            from PyQt5.QtWidgets import QApplication
            app = QApplication.instance()
            if app is not None:
                app.setStyleSheet(app.styleSheet())
        except Exception:
            pass

        # Re-apply per-widget custom stylesheets so they pick up the new theme
        self._refresh_themed_widgets()
        if hasattr(self, "batch_tab"):
            try:
                self.batch_tab.refresh_theme()
            except Exception:
                pass

        # Repaint custom widgets that depend on theme
        self.grid_preview.update()
        if hasattr(self, "thumb_stack"):
            for card in self.thumb_stack.cards:
                card.update()
        self._rebuild_tree()
        self.update()
        self.repaint()

    def _show_about(self):
        QMessageBox.about(
            self, "About Excel Image Inserter",
            f"<h3>Excel Image Inserter</h3>"
            f"<p>Version {APP_VERSION} (build {BUILD_NUMBER})</p>"
            f"<p>Created by I.Moskvin using Claude Opus 4.6</p>"
            f"<p>Batch insert images into Excel .xlsx files<br>"
            f"with grouping, TOC, and layout control.</p>"
        )

    def _on_settings_changed(self, *_):
        crop_key = self.combo_crop.currentText()
        crop = CROP_PRESETS.get(crop_key)
        start_col = self.le_start_col.text().strip().upper() or "A"
        self.grid_preview.update_params(
            groups=self.groups,
            cols=self.spin_cols.value(),
            crop_ratio=crop,
            start_col=start_col,
            start_row=self.spin_start_row.value(),
            placement="in_cell" if self.combo_placement.currentIndex() == 1 else "over",
            use_groups=self.cb_use_groups.isChecked(),
        )

    def _on_display_mode_changed(self, index):
        # Per image (0): show W/H toggle, hide aspect dropdown
        # Fixed ratio (1): show W/H toggle AND aspect dropdown
        # Manual (2): hide both, both spinboxes enabled
        if index == 0:
            self.lbl_anchor_axis.show()
            self.rb_anchor_w.show()
            self.rb_anchor_h.show()
            self.lbl_fixed_aspect.hide()
            self.combo_fixed_aspect.hide()
            self.spin_aspect_w.hide()
            self.spin_aspect_h.hide()
            self._apply_anchor_axis_enable()
        elif index == 1:
            self.lbl_anchor_axis.show()
            self.rb_anchor_w.show()
            self.rb_anchor_h.show()
            self.lbl_fixed_aspect.show()
            self.combo_fixed_aspect.show()
            is_custom = self.combo_fixed_aspect.currentText() == "Custom..."
            self.spin_aspect_w.setVisible(is_custom)
            self.spin_aspect_h.setVisible(is_custom)
            # Keep _display_aspect in sync with the active fixed ratio
            aw, ah = self._current_fixed_aspect()
            if aw and ah:
                self._display_aspect = aw / ah
            self._apply_anchor_axis_enable()
            # Recompute the dependent spinbox from the current anchor
            self._sync_fixed_spinbox()
        else:
            self.lbl_anchor_axis.hide()
            self.rb_anchor_w.hide()
            self.rb_anchor_h.hide()
            self.lbl_fixed_aspect.hide()
            self.combo_fixed_aspect.hide()
            self.spin_aspect_w.hide()
            self.spin_aspect_h.hide()
            self.spin_cm_w.setEnabled(True)
            self.spin_cm_h.setEnabled(True)

    def _current_fixed_aspect(self):
        """Return (w, h) tuple for the active aspect dropdown selection."""
        text = self.combo_fixed_aspect.currentText()
        if text == "Custom...":
            return self.spin_aspect_w.value(), self.spin_aspect_h.value()
        try:
            w, h = text.split(":")
            return int(w), int(h)
        except (ValueError, AttributeError):
            return 4, 3

    def _apply_anchor_axis_enable(self):
        """Enable/disable W and H spinboxes based on selected anchor axis.
        Only applies in Per image / Fixed ratio modes (caller responsible)."""
        if self.rb_anchor_w.isChecked():
            self.spin_cm_w.setEnabled(True)
            self.spin_cm_h.setEnabled(False)
        else:
            self.spin_cm_w.setEnabled(False)
            self.spin_cm_h.setEnabled(True)

    def _on_anchor_axis_changed(self, _checked=False):
        mode = self.combo_display_mode.currentIndex()
        if mode == 2:
            return
        self._apply_anchor_axis_enable()
        if mode == 1:
            self._sync_fixed_spinbox()

    def _on_aspect_changed(self, *_):
        # Toggle custom spinboxes visibility based on dropdown
        is_custom = self.combo_fixed_aspect.currentText() == "Custom..."
        if self.combo_display_mode.currentIndex() == 1:
            self.spin_aspect_w.setVisible(is_custom)
            self.spin_aspect_h.setVisible(is_custom)
        aw, ah = self._current_fixed_aspect()
        if aw and ah:
            self._display_aspect = aw / ah
        if self.combo_display_mode.currentIndex() == 1:
            self._sync_fixed_spinbox()

    def _sync_fixed_spinbox(self):
        """In Fixed-ratio mode, drive the disabled spinbox from the active
        one through the current aspect ratio."""
        if self._cm_updating:
            return
        aspect = max(self._display_aspect, 0.01)
        self._cm_updating = True
        try:
            if self.rb_anchor_w.isChecked():
                self.spin_cm_h.setValue(self.spin_cm_w.value() / aspect)
            else:
                self.spin_cm_w.setValue(self.spin_cm_h.value() * aspect)
        finally:
            self._cm_updating = False

    def _on_cm_w_changed(self, val):
        if self._cm_updating:
            return
        if self.combo_display_mode.currentIndex() == 1 and self.rb_anchor_w.isChecked():
            self._cm_updating = True
            self.spin_cm_h.setValue(val / max(self._display_aspect, 0.01))
            self._cm_updating = False

    def _on_cm_h_changed(self, val):
        if self._cm_updating:
            return
        if self.combo_display_mode.currentIndex() == 1 and self.rb_anchor_h.isChecked():
            self._cm_updating = True
            self.spin_cm_w.setValue(val * self._display_aspect)
            self._cm_updating = False

    # ── Insert ────────────────────────────────────────────────────────────
    def _do_insert(self):
        file_path = self.le_file.text().strip()
        if self.rb_open.isChecked() and (not file_path or not os.path.exists(file_path)):
            QMessageBox.warning(self, "Error", "Please select an existing Excel file.")
            return
        if self.rb_new.isChecked() and not file_path:
            QMessageBox.warning(self, "Error", "Please specify a file path to save.")
            return
        if not self.image_paths:
            QMessageBox.warning(self, "Error", "No images to insert.")
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

        # Determine insert position
        insert_after_name = None
        if sheet_new and self.combo_insert_after.isVisible():
            sel = self.combo_insert_after.currentIndex()
            if sel > 0:
                insert_after_name = self.combo_insert_after.currentText()

        params = {
            "excel_path": file_path if self.rb_open.isChecked() else None,
            "save_path": file_path,
            "sheet_new": sheet_new,
            "sheet_name": sheet_name,
            "insert_after_name": insert_after_name,
            "groups": [dict(g) for g in self.groups],
            "resize_px_w": self._get_resize_px()[0],
            "resize_px_h": self._get_resize_px()[1],
            "display_w_cm": self.spin_cm_w.value(),
            "display_h_cm": self.spin_cm_h.value(),
            "display_mode": self.combo_display_mode.currentIndex(),
            "anchor_axis": "W" if self.rb_anchor_w.isChecked() else "H",
            "fixed_aspect": self._current_fixed_aspect(),
            "crop_ratio": crop,
            "grid_cols": self.spin_cols.value(),
            "start_col": start_col,
            "start_row": self.spin_start_row.value(),
            "placement": "in_cell" if self.combo_placement.currentIndex() == 1 else "over",
            "gap_h_cm": self.spin_gap_h.value(),
            "gap_v_cm": self.spin_gap_v.value(),
            "create_toc": self.cb_toc.isChecked() and self.cb_use_groups.isChecked(),
            "use_groups": self.cb_use_groups.isChecked(),
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
