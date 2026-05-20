#!/usr/bin/env python3
"""Excel Image Inserter — PyQt5 utility for batch-inserting images into Excel."""

import sys

from PyQt5.QtWidgets import QApplication

from app.ui.main_window import MainWindow


def main():
    app = QApplication(sys.argv)
    app.setStyle("Fusion")
    win = MainWindow()
    screen = app.primaryScreen().availableGeometry()
    win.resize(win.minimumSizeHint().width(), int(screen.height() * 0.9))
    win.move(screen.x(), screen.y())
    win.show()
    sys.exit(app.exec_())


if __name__ == "__main__":
    main()
