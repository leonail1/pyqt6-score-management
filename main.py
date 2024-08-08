import os
import sys

sys.path.append(os.getcwd())

from PyQt6.QtWidgets import QApplication

from my_window.MainWindow import MainWindow

if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = MainWindow()
    window.show()
    sys.exit(app.exec())