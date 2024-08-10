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

from file_import.action_creator import ActionCreator
from file_import.table_file_dealer import FileDealer
from file_import.menu_manager import MenuManager


class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("学生成绩管理系统")
        self.resize(QSize(500, 300))
        self.menu_manager = MenuManager(self)
        self.file_dealer = self.menu_manager.file_dealer  # 获取 FileDealer 实例
        self.setup_central_widget()
        self.menu_manager.setup_menu()  # 菜单栏
        self.setup_statusbar()  # 状态栏

    def setup_central_widget(self):
        """
        欢迎界面
        """
        central_widget = QWidget()
        layout = QVBoxLayout()

        welcome_label = QLabel("欢迎使用学生成绩管理系统")
        welcome_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        welcome_label.setStyleSheet("font-size: 24pt; margin-bottom: 20px;")
        layout.addWidget(welcome_label)

        # 创建学号输入框
        input_layout = QHBoxLayout()
        student_id_label = QLabel("请输入学号：")
        self.student_id_input = QLineEdit()
        self.student_id_input.setPlaceholderText("学号必须为14位纯数字")
        self.student_id_input.setMaxLength(14)  # 设置最大长度为14
        self.student_id_input.setValidator(StudentIDValidator())  # 只允许输入14位数字
        self.file_dealer.set_default_student_id(self.student_id_input)  # 设置默认学号
        self.student_id_input.returnPressed.connect(lambda: self.file_dealer.process_student_id(self.student_id_input))

        # 创建确认按钮
        confirm_button = QPushButton("确认")
        confirm_button.clicked.connect(lambda: self.file_dealer.process_student_id(self.student_id_input))

        input_layout.addWidget(student_id_label)
        input_layout.addWidget(self.student_id_input)
        input_layout.addWidget(confirm_button)
        layout.addLayout(input_layout)

        central_widget.setLayout(layout)
        self.setCentralWidget(central_widget)

    def setup_statusbar(self):
        # 设置状态栏
        self.setStatusBar(QStatusBar(self))


class StudentIDValidator(QValidator):
    def validate(self, input_string, pos):
        """
        对输入的学号进行验证，要求14位纯数字
        """
        if input_string.isdigit() and len(input_string) <= 14:
            return QValidator.State.Acceptable, input_string, pos
        elif input_string == "":
            return QValidator.State.Intermediate, input_string, pos
        else:
            return QValidator.State.Invalid, input_string, pos
