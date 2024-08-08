import os
import sys

sys.path.append(os.getcwd())

from PyQt6.QtWidgets import (
    QDialog, QVBoxLayout, QTableView, QLabel,
    QHeaderView, QScrollArea, QWidget, QApplication
)
from PyQt6.QtCore import Qt
from PyQt6.QtGui import QStandardItemModel, QStandardItem

from my_code.student_score_analyzer import StudentScoreAnalyzer

class StudentInfoWindow(QDialog):
    def __init__(self, chosen_window: str, **kwargs):
        super().__init__()

        self.student_score_analyzer = StudentScoreAnalyzer(self)

        self.setWindowTitle("学生信息")
        self.resize(1400, 800)
        self.setup_score_list_view_ui(**kwargs)

    def setup_score_list_view_ui(self, student_id: str):
        score_data = self.student_score_analyzer.load_score_data(student_id=student_id)

        if score_data is None:
            error_label = QLabel("无法加载学生数据")
            self.setLayout(QVBoxLayout())
            self.layout().addWidget(error_label)
            return

        main_layout = QVBoxLayout()
        student_info = score_data.get("student_info", {})
        info_label = QLabel(f"姓名: {student_info.get('姓名', 'N/A')}  学号: {student_info.get('学号', 'N/A')}")
        main_layout.addWidget(info_label)

        self.model = QStandardItemModel()

        headers = ["课程名", "课程性质", "学分", "学年学期", "等级成绩", "绩点"]
        self.model.setHorizontalHeaderLabels(headers)

        scores = score_data.get("scores", [])
        for score in scores:
            row_items = [
                QStandardItem(str(score.get("课程名", ""))),
                QStandardItem(str(score.get("课程性质", ""))),
                QStandardItem(str(score.get("学分", ""))),
                QStandardItem(str(score.get("学年学期", ""))),
                QStandardItem(str(score.get("等级成绩", ""))),
                QStandardItem(str(score.get("绩点", "")))
            ]
            self.model.appendRow(row_items)

        self.table = QTableView()
        self.table.setModel(self.model)
        self.table.setSortingEnabled(True)
        self.table.setEditTriggers(QTableView.EditTrigger.NoEditTriggers)

        header = self.table.horizontalHeader()
        header.setSectionResizeMode(QHeaderView.ResizeMode.Stretch)

        main_layout.addWidget(self.table)

        container = QWidget()
        container.setLayout(main_layout)

        scroll_area = QScrollArea()
        scroll_area.setWidgetResizable(True)
        scroll_area.setWidget(container)

        self.setLayout(QVBoxLayout())
        self.layout().addWidget(scroll_area)