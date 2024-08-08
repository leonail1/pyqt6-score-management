import os
import sys

sys.path.append(os.getcwd())

from PyQt6.QtWidgets import (
    QDialog, QVBoxLayout, QHBoxLayout, QTableView, QLabel,
    QHeaderView, QScrollArea, QWidget, QApplication, QPushButton, QListWidget, QCheckBox, QListWidgetItem
)
from PyQt6.QtCore import Qt, QSortFilterProxyModel
from PyQt6.QtGui import QStandardItemModel, QStandardItem, QFont

from my_code.student_score_analyzer import StudentScoreAnalyzer


class CustomSortFilterProxyModel(QSortFilterProxyModel):
    def __init__(self):
        super().__init__()
        self.column_filters = {}

        # 设置全局字体
        font = QFont()
        font.setPointSize(16)  # 设置字体大小
        font.setBold(True)  # 设置字体粗细

    def setColumnFilter(self, column, values):
        self.column_filters[column] = values
        self.invalidateFilter()

    def filterAcceptsRow(self, source_row, source_parent):
        for column, values in self.column_filters.items():
            index = self.sourceModel().index(source_row, column, source_parent)
            if index.isValid():
                if self.sourceModel().data(index) not in values and values:
                    return False
            else:
                return False
        return True


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
        self.proxy_model = CustomSortFilterProxyModel()
        self.proxy_model.setSourceModel(self.model)

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
            # 设置每个单元格的对齐方式
            for item in row_items:
                item.setTextAlignment(Qt.AlignmentFlag.AlignCenter)
            self.model.appendRow(row_items)

        filter_layout = QHBoxLayout()
        self.filter_buttons = []
        for column in range(len(headers)):
            button = QPushButton(f"过滤 {headers[column]}")
            button.clicked.connect(lambda checked, col=column: self.show_filter_dialog(col))
            filter_layout.addWidget(button)
            self.filter_buttons.append(button)

        main_layout.addLayout(filter_layout)

        self.table = QTableView()
        self.table.setModel(self.proxy_model)
        self.table.setSortingEnabled(True)
        self.table.setEditTriggers(QTableView.EditTrigger.NoEditTriggers)

        header = self.table.horizontalHeader()
        header.setSectionResizeMode(QHeaderView.ResizeMode.Stretch)

        # 设置表头的对齐方式
        for i in range(self.model.columnCount()):
            self.table.horizontalHeader().setDefaultAlignment(Qt.AlignmentFlag.AlignCenter)

        main_layout.addWidget(self.table)

        container = QWidget()
        container.setLayout(main_layout)

        scroll_area = QScrollArea()
        scroll_area.setWidgetResizable(True)
        scroll_area.setWidget(container)

        self.setLayout(QVBoxLayout())
        self.layout().addWidget(scroll_area)

    def show_filter_dialog(self, column):
        dialog = QDialog(self)
        dialog.setWindowTitle(f"过滤 {self.model.headerData(column, Qt.Orientation.Horizontal)}")
        layout = QVBoxLayout()

        list_widget = QListWidget()
        unique_values = sorted(set(self.model.item(row, column).text() for row in range(self.model.rowCount()) if
                                   self.model.item(row, column) is not None))

        select_all_checkbox = QCheckBox("全选/反选")
        select_all_checkbox.setChecked(True)
        select_all_checkbox.stateChanged.connect(lambda state: self.toggle_all_items(list_widget, state))
        layout.addWidget(select_all_checkbox)

        for value in unique_values:
            item = QListWidgetItem(value)
            item.setFlags(item.flags() | Qt.ItemFlag.ItemIsUserCheckable)
            item.setCheckState(Qt.CheckState.Checked)
            list_widget.addItem(item)

        layout.addWidget(list_widget)

        apply_button = QPushButton("应用")
        apply_button.clicked.connect(lambda: self.apply_filter(column, list_widget, dialog))
        layout.addWidget(apply_button)

        dialog.setLayout(layout)
        dialog.exec()

    def toggle_all_items(self, list_widget, state):
        for i in range(list_widget.count()):
            item = list_widget.item(i)
            if item:
                item.setCheckState(Qt.CheckState.Checked if state == Qt.CheckState.Checked else Qt.CheckState.Unchecked)

    def apply_filter(self, column, list_widget, dialog):
        selected_values = [list_widget.item(i).text() for i in range(list_widget.count())
                           if list_widget.item(i) and list_widget.item(i).checkState() == Qt.CheckState.Checked]
        self.proxy_model.setColumnFilter(column, selected_values)
        dialog.accept()
