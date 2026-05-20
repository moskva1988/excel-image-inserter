#!/usr/bin/env python3
"""Excel Image Inserter — PyQt5 utility for batch-inserting images into Excel."""

import sys

from PyQt5.QtCore import Qt, QSettings
from PyQt5.QtWidgets import QApplication

# HiDPI must be set BEFORE QApplication is constructed
QApplication.setAttribute(Qt.AA_EnableHighDpiScaling, True)
QApplication.setAttribute(Qt.AA_UseHighDpiPixmaps, True)

from qfluentwidgets import setTheme, Theme  # noqa: E402

from app.ui.main_window import MainWindow  # noqa: E402


def _load_initial_theme():
    """Read persisted theme preference (System/Light/Dark) and apply it."""
    settings = QSettings("ExcelImageInserter", "ExcelImageInserter")
    pref = settings.value("ui/theme", "System")
    mapping = {"Light": Theme.LIGHT, "Dark": Theme.DARK, "System": Theme.AUTO}
    setTheme(mapping.get(pref, Theme.AUTO))


def main():
    app = QApplication(sys.argv)
    _load_initial_theme()
    win = MainWindow()
    screen = app.primaryScreen().availableGeometry()
    win.resize(win.minimumSizeHint().width(), int(screen.height() * 0.9))
    win.move(screen.x(), screen.y())
    win.show()
    sys.exit(app.exec_())


if __name__ == "__main__":
    main()
