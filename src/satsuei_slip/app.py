from __future__ import annotations

import sys
from pathlib import Path

from PySide6.QtWidgets import QApplication

try:
    from satsuei_slip.main_window import MainWindow
except ModuleNotFoundError:
    src_dir = Path(__file__).resolve().parents[1]
    if str(src_dir) not in sys.path:
        sys.path.insert(0, str(src_dir))
    from satsuei_slip.main_window import MainWindow


def main() -> int:
    app = QApplication(sys.argv)
    app.setApplicationName("SatsueiSlip")
    app.setOrganizationName("SatsueiSlip")

    window = MainWindow()
    window.show()

    return app.exec()
