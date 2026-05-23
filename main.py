#!/usr/bin/env python3
"""Excel Image Inserter — PyQt5 utility for batch-inserting images into Excel."""

import os
import sys
import tempfile
import traceback as _tb

# Debug logging: redirect stdout/stderr to file in OS temp dir + capture unhandled exceptions
_LOG_PATH = os.path.join(tempfile.gettempdir(), "excel-image-inserter.log")
try:
    _log_fh = open(_LOG_PATH, "a", encoding="utf-8", buffering=1)
    _log_fh.write(f"\n{'='*60}\nApp start: {__import__('datetime').datetime.now().isoformat()}\nLog: {_LOG_PATH}\n{'='*60}\n")
    sys.stdout = _log_fh
    sys.stderr = _log_fh
except Exception:
    pass

def _excepthook(exc_type, exc_value, exc_tb):
    msg = "".join(_tb.format_exception(exc_type, exc_value, exc_tb))
    try:
        sys.stderr.write(f"\nUNHANDLED EXCEPTION:\n{msg}\n")
        sys.stderr.flush()
    except Exception:
        pass
sys.excepthook = _excepthook

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
