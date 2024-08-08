import os
import sys

sys.path.append(os.getcwd())

from PyQt6.QtWidgets import (
    QDialog, QVBoxLayout, QHBoxLayout, QTableView, QLabel,
    QHeaderView, QScrollArea, QWidget, QApplication, QPushButton, QListWidget, QCheckBox, QListWidgetItem, QMessageBox
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
                if not values:  # 如果过滤列表为空，不显示任何行
                    return False
                if self.sourceModel().data(index) not in values:
                    return False
            else:
                return False
        return True


class StudentInfoWindow(QDialog):
    def __init__(self, chosen_window: str, student_id):
        super().__init__()

        self.data_modified = False
        self.score_data = None
        self.student_score_analyzer = StudentScoreAnalyzer(self)
        # 添加一部字典来存储每列的选中状态
        self.column_filter_states = {}
        self.student_id = student_id

        self.setWindowTitle("学生信息")
        self.resize(1400, 800)
        self.setup_score_list_view_ui(self.student_id)

    def setup_score_list_view_ui(self, student_id: str):
        self.score_data = self.student_score_analyzer.load_score_data(student_id=student_id)

        if self.score_data is None:
            error_label = QLabel("无法加载学生数据")
            self.setLayout(QVBoxLayout())
            self.layout().addWidget(error_label)
            return

        main_layout = QVBoxLayout()
        student_info = self.score_data.get("student_info", {})
        info_label = QLabel(f"姓名: {student_info.get('姓名', 'N/A')}  学号: {student_info.get('学号', 'N/A')}")
        main_layout.addWidget(info_label)

        self.model = QStandardItemModel()
        self.proxy_model = CustomSortFilterProxyModel()
        self.proxy_model.setSourceModel(self.model)

        headers = ["课程名", "课程性质", "学分", "学年学期", "等级成绩", "绩点"]
        self.model.setHorizontalHeaderLabels(headers)

        scores = self.score_data.get("scores", [])
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

        # 设置表格为可编辑
        self.table.setEditTriggers(QTableView.EditTrigger.DoubleClicked | QTableView.EditTrigger.EditKeyPressed)
        # 连接数据更改信号到处理函数
        self.model.dataChanged.connect(self.on_table_view_data_changed)

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

    def on_table_view_data_changed(self, top_left, bottom_right, roles):
        for row in range(top_left.row(), bottom_right.row() + 1):
            for column in range(top_left.column(), bottom_right.column() + 1):
                index = self.model.index(row, column)
                item = self.model.itemFromIndex(index)
                if item:
                    column_name = self.model.headerData(column, Qt.Orientation.Horizontal)
                    value = item.text()

                    # 更新 score_data 中的值
                    if 0 <= row < len(self.score_data["scores"]):
                        self.score_data["scores"][row][column_name] = value
                        print(f"Data updated in score_data: row {row}, column {column_name}, value {value}")
                    else:
                        print(f"Failed to update data: row {row} out of range")

        # 设置一个标志，表示数据已被修改
        self.data_modified = True

    def show_filter_dialog(self, column):
        dialog = QDialog(self)
        dialog.setWindowTitle(f"过滤 {self.model.headerData(column, Qt.Orientation.Horizontal)}")
        layout = QVBoxLayout()

        # 创建按钮布局
        button_layout = QHBoxLayout()
        select_all_button = QPushButton("全选")
        select_none_button = QPushButton("全不选")
        button_layout.addWidget(select_all_button)
        button_layout.addWidget(select_none_button)
        layout.addLayout(button_layout)

        list_widget = QListWidget()
        unique_values = sorted(set(self.model.item(row, column).text() for row in range(self.model.rowCount()) if
                                   self.model.item(row, column) is not None))

        # 如果这个列还没有保存的状态，初始化为全选
        if column not in self.column_filter_states:
            self.column_filter_states[column] = {value: True for value in unique_values}

        # 添加所有选项到列表中，根据保存的状态设置选中状态
        for value in unique_values:
            item = QListWidgetItem(value)
            item.setFlags(item.flags() | Qt.ItemFlag.ItemIsUserCheckable)
            item.setCheckState(
                Qt.CheckState.Checked if self.column_filter_states[column][value] else Qt.CheckState.Unchecked)
            list_widget.addItem(item)

        # 连接全选按钮
        select_all_button.clicked.connect(lambda: self.set_all_items(list_widget, True, column))

        # 连接全不选按钮
        select_none_button.clicked.connect(lambda: self.set_all_items(list_widget, False, column))

        # 连接itemChanged信号到更新字典的函数
        list_widget.itemChanged.connect(lambda item: self.update_filter_state(column, item))

        layout.addWidget(list_widget)

        apply_button = QPushButton("应用")
        apply_button.clicked.connect(lambda: self.apply_filter(column, list_widget, dialog))
        layout.addWidget(apply_button)

        dialog.setLayout(layout)
        dialog.exec()

    def set_all_items(self, list_widget, checked, column):
        for i in range(list_widget.count()):
            item = list_widget.item(i)
            item.setCheckState(Qt.CheckState.Checked if checked else Qt.CheckState.Unchecked)
            self.column_filter_states[column][item.text()] = checked

    def update_filter_state(self, column, item):
        # 更新字典中的状态
        self.column_filter_states[column][item.text()] = (item.checkState() == Qt.CheckState.Checked)

    def apply_filter(self, column, list_widget, dialog):
        # 获取所有选中的值
        selected_values = [list_widget.item(i).text() for i in range(list_widget.count())
                           if list_widget.item(i).checkState() == Qt.CheckState.Checked]

        # 应用过滤器
        self.proxy_model.setColumnFilter(column, selected_values)

        dialog.accept()

    def closeEvent(self, event):
        if self.data_modified:
            reply = QMessageBox.question(
                self,
                '保存更改',
                "数据已被修改。是否保存更改？",
                QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No | QMessageBox.StandardButton.Cancel
            )

            if reply == QMessageBox.StandardButton.Yes:
                if self.student_score_analyzer.save_score_data(student_id=self.student_id, score_data=self.score_data):
                    event.accept()  # 保存成功，允许关闭窗口
                else:
                    QMessageBox.warning(self, "保存失败", "保存数据时发生错误。")
                    event.ignore()  # 保存失败，不关闭窗口
            elif reply == QMessageBox.StandardButton.No:
                event.accept()  # 不保存，允许关闭窗口
            else:  # Cancel
                event.ignore()  # 取消关闭操作
        else:
            event.accept()  # 如果数据没有被修改，直接关闭窗口
