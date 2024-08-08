import sys
import os

sys.path.append(os.getcwd())

import json
from PyQt6.QtWidgets import QMainWindow, QLabel, QStatusBar, QMenu, QFileDialog, QMessageBox, QApplication, QListWidget, \
    QPushButton, QVBoxLayout, QWidget, QHBoxLayout, QLineEdit
from PyQt6.QtGui import QIcon, QIntValidator, QValidator
from PyQt6.QtCore import Qt, QSize
from PyQt6.QtGui import QAction
import pandas as pd

from my_window.MainWindow import MainWindow

if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = MainWindow()
    window.show()
    sys.exit(app.exec())