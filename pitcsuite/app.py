import sys

from PySide6.QtWidgets import QApplication

from pitcsuite.icons import app_icon
from pitcsuite.main_window import MainWindow
from pitcsuite.ui_helpers import DARK_STYLE


def main():
    app = QApplication(sys.argv)
    app.setWindowIcon(app_icon())
    app.setStyleSheet(DARK_STYLE)
    win = MainWindow()
    win.show()
    return app.exec()


if __name__ == "__main__":
    raise SystemExit(main())
