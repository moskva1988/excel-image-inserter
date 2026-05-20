from pathlib import Path

from PyQt5.QtWidgets import (
    QWidget, QScrollArea, QVBoxLayout, QHBoxLayout,
)
from PyQt5.QtCore import Qt, QPoint, pyqtSignal
from PyQt5.QtGui import QPixmap, QPainter, QPen, QColor, QFont, QBrush

from app.core.image_processor import estimate_size
from app.core.models import THUMB_SIZE

try:
    from qfluentwidgets import isDarkTheme, themeColor
except Exception:  # pragma: no cover
    def isDarkTheme():
        return False

    def themeColor():
        return QColor("#6366f1")


# ── Thumbnail stack widget ─────────────────────────────────────────────────────
class ThumbCard(QWidget):
    delete_requested = pyqtSignal(str)
    selection_toggled = pyqtSignal(str, bool)

    def __init__(self, path, orig_mb, est_mb, w, h):
        super().__init__()
        self.path = path
        self.orig_mb = orig_mb
        self.est_mb = est_mb
        self.img_w = w
        self.img_h = h
        self.selected = False
        self._drag_start = None
        self.pixmap = QPixmap(path).scaled(THUMB_SIZE, THUMB_SIZE, Qt.KeepAspectRatio, Qt.SmoothTransformation)
        self.setFixedSize(self.pixmap.width(), self.pixmap.height())
        self.setToolTip(f"{Path(path).name}\n{w}x{h}\n{orig_mb:.2f} MB -> {est_mb:.2f} MB")

    def paintEvent(self, event):
        p = QPainter(self)
        p.setRenderHint(QPainter.Antialiasing)
        dark = isDarkTheme()
        accent = themeColor()
        p.drawPixmap(0, 0, self.pixmap)
        if self.selected:
            p.setPen(QPen(accent, 3))
            p.setBrush(Qt.NoBrush)
            p.drawRect(1, 1, self.width() - 2, self.height() - 2)
        bar_h = 18
        bar_y = self.height() - bar_h
        if dark:
            bar_bg = QColor(20, 20, 20, 220)
            orig_col = QColor("#e0e0e0")
            est_col = QColor("#5fd87f")
            btn_bg = QColor(60, 60, 60, 230)
            btn_fg = QColor("#f0f0f0")
        else:
            bar_bg = QColor(255, 255, 255, 200)
            orig_col = QColor("#333")
            est_col = QColor("#16a34a")
            btn_bg = QColor(255, 255, 255, 220)
            btn_fg = QColor("#333")
        p.fillRect(0, bar_y, self.width(), bar_h, bar_bg)
        p.setFont(QFont("Arial", 8))
        p.setPen(orig_col)
        p.drawText(4, bar_y, self.width() // 2, bar_h, Qt.AlignLeft | Qt.AlignVCenter, f"{self.orig_mb:.2f}MB")
        p.setPen(est_col)
        p.drawText(self.width() // 2, bar_y, self.width() // 2 - 4, bar_h, Qt.AlignRight | Qt.AlignVCenter, f"{self.est_mb:.2f}MB")
        btn_r = 9
        cx = self.width() - btn_r - 4
        cy = btn_r + 4
        p.setBrush(btn_bg)
        p.setPen(Qt.NoPen)
        p.drawEllipse(QPoint(cx, cy), btn_r, btn_r)
        p.setPen(btn_fg)
        p.setFont(QFont("Arial", 9, QFont.Bold))
        p.drawText(cx - btn_r, cy - btn_r, btn_r * 2, btn_r * 2, Qt.AlignCenter, "×")
        p.end()

    def mousePressEvent(self, event):
        btn_r = 9
        cx = self.width() - btn_r - 4
        cy = btn_r + 4
        if (event.pos().x() - cx) ** 2 + (event.pos().y() - cy) ** 2 <= (btn_r + 3) ** 2:
            self.delete_requested.emit(self.path)
            return
        self._drag_start = event.pos()

    def mouseMoveEvent(self, event):
        if self._drag_start and (event.pos() - self._drag_start).manhattanLength() > 10:
            from PyQt5.QtCore import QMimeData
            from PyQt5.QtGui import QDrag
            drag = QDrag(self)
            mime = QMimeData()
            mime.setText(self.path)
            drag.setMimeData(mime)
            drag.setPixmap(self.pixmap.scaled(60, 60, Qt.KeepAspectRatio, Qt.SmoothTransformation))
            drag.exec_(Qt.MoveAction)
            self._drag_start = None

    def mouseReleaseEvent(self, event):
        if self._drag_start:
            self.selected = not self.selected
            self.selection_toggled.emit(self.path, self.selected)
            self.update()
        self._drag_start = None


class ThumbStackView(QScrollArea):
    delete_requested = pyqtSignal(str)
    order_changed = pyqtSignal(list)

    def __init__(self):
        super().__init__()
        self.setWidgetResizable(True)
        self.setAcceptDrops(True)
        self.setStyleSheet("ThumbStackView { border: 1px solid palette(mid); border-radius: 6px; }")
        self.container = QWidget()
        self.container.setAcceptDrops(True)
        self.flow = FlowLayout(self.container)
        self.flow.setSpacing(8)
        self.setWidget(self.container)
        self.cards = []
        self.selected_paths = set()
        self._paths = []

    def set_images(self, paths, max_w, max_h):
        self.flow.clear_widgets()
        for c in self.cards:
            c.setParent(None)
            c.deleteLater()
        self.cards.clear()
        self.selected_paths.clear()
        self._paths = list(paths)
        for path in paths:
            orig_mb, est_mb, w, h = estimate_size(path, max_w, max_h)
            card = ThumbCard(path, orig_mb, est_mb, w, h)
            card.delete_requested.connect(self._on_delete)
            card.selection_toggled.connect(self._on_selection)
            self.cards.append(card)
        self.flow.set_widgets(self.cards)

    def _on_delete(self, path):
        self.delete_requested.emit(path)

    def _on_selection(self, path, selected):
        if selected:
            self.selected_paths.add(path)
        else:
            self.selected_paths.discard(path)

    def get_selected(self):
        return list(self.selected_paths)

    def dragEnterEvent(self, event):
        if event.mimeData().hasText():
            event.acceptProposedAction()

    def dragMoveEvent(self, event):
        event.acceptProposedAction()

    def dropEvent(self, event):
        src_path = event.mimeData().text()
        if src_path not in self._paths:
            return
        drop_pos = self.container.mapFrom(self, event.pos())
        target_idx = len(self._paths) - 1
        for i, card in enumerate(self.cards):
            if card.geometry().contains(drop_pos):
                target_idx = i
                break
        src_idx = self._paths.index(src_path)
        if src_idx == target_idx:
            return
        self._paths.pop(src_idx)
        self._paths.insert(target_idx, src_path)
        self.order_changed.emit(list(self._paths))
        event.acceptProposedAction()


class FlowLayout(QVBoxLayout):
    def __init__(self, parent=None):
        super().__init__(parent)
        self._widgets = []

    def clear_widgets(self):
        self._widgets.clear()
        while self.count():
            item = self.takeAt(0)
            if item.layout():
                while item.layout().count():
                    item.layout().takeAt(0)

    def set_widgets(self, widgets):
        self._widgets = list(widgets)
        self._relayout()

    def addWidget(self, widget):
        self._widgets.append(widget)
        self._relayout()

    def _relayout(self):
        while self.count():
            item = self.takeAt(0)
            if item.layout():
                while item.layout().count():
                    item.layout().takeAt(0)
        if not self._widgets:
            return
        parent = self.parentWidget()
        available_w = parent.width() if parent else 600
        card_w = THUMB_SIZE + 10
        cols = max(1, available_w // card_w)
        row_lay = None
        for i, w in enumerate(self._widgets):
            if i % cols == 0:
                row_lay = QHBoxLayout()
                row_lay.setAlignment(Qt.AlignLeft)
                super().addLayout(row_lay)
            row_lay.addWidget(w)
