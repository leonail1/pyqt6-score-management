import os
import sys

sys.path.append(os.getcwd())

from PyQt6.QtWidgets import (
    QDialog, QVBoxLayout, QHBoxLayout, QTableView, QLabel,
    QHeaderView, QScrollArea, QWidget, QApplication, QPushButton, QListWidget, QCheckBox, QListWidgetItem, QMessageBox
)
from PyQt6.QtCore import Qt, QSortFilterProxyModel
from PyQt6.QtGui import QStandardItemModel, QStandardItem, QFont

from file_import.student_score_analyzer import StudentScoreAnalyzer
from .DegreeProgressShow import create_degree_progress_window


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
    def __init__(self, student_id):
        super().__init__()

        self.data_modified = False
        self.score_data = None
        self.student_score_analyzer = StudentScoreAnalyzer(self)
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

        scores = self.score_data.get("scores", [])
        if scores:
            # 动态获取所有列名
            headers = list(scores[0].keys())
            self.model.setHorizontalHeaderLabels(headers)

            for score in scores:
                row_items = [QStandardItem(str(score.get(header, ""))) for header in headers]
                # 设置每个单元格的对齐方式
                for item in row_items:
                    item.setTextAlignment(Qt.AlignmentFlag.AlignCenter)
                self.model.appendRow(row_items)

            filter_layout = QHBoxLayout()
            self.filter_buttons = []
            for column, header in enumerate(headers):
                button = QPushButton(f"过滤 {header}")
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
            header.setSectionResizeMode(QHeaderView.ResizeMode.ResizeToContents)  # 自适应内容宽度
            header.setStretchLastSection(True)  # 确保最后一列填充剩余空间

            # 设置表头的对齐方式
            for i in range(self.model.columnCount()):
                self.table.horizontalHeader().setDefaultAlignment(Qt.AlignmentFlag.AlignCenter)

            main_layout.addWidget(self.table)

            # 添加加权绩点和加权分数的标签
            self.gpa_label = QLabel()
            self.weighted_score_label = QLabel()
            main_layout.addWidget(self.gpa_label)
            main_layout.addWidget(self.weighted_score_label)

            # 初始计算并显示加权绩点和加权分数
            self.update_weighted_calculations()
        else:
            error_label = QLabel("没有成绩数据")
            main_layout.addWidget(error_label)

        container = QWidget()
        container.setLayout(main_layout)

        scroll_area = QScrollArea()
        scroll_area.setWidgetResizable(True)
        scroll_area.setWidget(container)

        # 确保滚动区域可以横向和纵向滚动
        scroll_area.setHorizontalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAsNeeded)
        scroll_area.setVerticalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAsNeeded)

        # 创建按钮布局
        button_layout = QHBoxLayout()

        # 添加"显示学位进度"按钮
        show_progress_button = QPushButton("显示学位进度")
        show_progress_button.clicked.connect(self.show_degree_progress)
        button_layout.addWidget(show_progress_button)

        # 添加"关闭窗口"按钮
        close_button = QPushButton("关闭窗口")
        close_button.clicked.connect(self.close)
        close_button.setDefault(True)  # 设置为默认按钮
        button_layout.addWidget(close_button)

        # 创建主布局
        main_layout = QVBoxLayout()
        main_layout.addWidget(scroll_area)
        main_layout.addLayout(button_layout)

        self.setLayout(main_layout)

    def show_degree_progress(self):
        progress_window = create_degree_progress_window(student_id=self.student_id, parent=self)
        if progress_window:
            progress_window.show()

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

        # 更新加权计算结果
        self.update_weighted_calculations()

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
            self.column_filter_states[column] = {}

        # 添加所有选项到列表中，根据保存的状态设置选中状态
        for value in unique_values:
            item = QListWidgetItem(value)
            item.setFlags(item.flags() | Qt.ItemFlag.ItemIsUserCheckable)
            # 如果值不在字典中，默认设置为选中状态
            if value not in self.column_filter_states[column]:
                self.column_filter_states[column][value] = True
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

        # 更新加权计算结果
        self.update_weighted_calculations()

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

    def update_weighted_calculations(self):
        total_credits = 0
        total_gpa_points = 0
        total_score_points = 0

        # 找到对应的列索引
        credit_index = -1
        gpa_index = -1
        score_index = -1
        course_type_index = -1

        for col in range(self.model.columnCount()):
            header = self.model.headerData(col, Qt.Orientation.Horizontal)
            if header == "学分":
                credit_index = col
            elif header in ["GPA", "绩点"]:
                gpa_index = col
            elif header in ["总成绩", "学分成绩"]:
                score_index = col
            elif header == "课程性质":
                course_type_index = col

        if credit_index == -1:
            self.gpa_label.setText("加权绩点: 未找到学分列")
            self.weighted_score_label.setText("加权分数: 未找到学分列")
            return

        for row in range(self.proxy_model.rowCount()):
            # 检查课程性质
            if course_type_index != -1:
                course_type = self.proxy_model.data(self.proxy_model.index(row, course_type_index))
                if course_type == "校选":
                    continue  # 跳过校选课程

            credit = float(self.proxy_model.data(self.proxy_model.index(row, credit_index)) or 0)

            if score_index != -1:
                score = self.proxy_model.data(self.proxy_model.index(row, score_index))
                if score == "合格":
                    continue  # 跳过成绩为"合格"的课程
                try:
                    score = float(score or 0)
                    total_score_points += score * credit
                    total_credits += credit
                except ValueError:
                    # 如果成绩无法转换为浮点数，跳过这门课程
                    continue

            if gpa_index != -1:
                gpa = self.proxy_model.data(self.proxy_model.index(row, gpa_index))
                try:
                    gpa = float(gpa or 0)
                    total_gpa_points += gpa * credit
                except ValueError:
                    # 如果GPA无法转换为浮点数，跳过GPA计算
                    pass

        if total_credits > 0:
            weighted_gpa = total_gpa_points / total_credits if gpa_index != -1 else None
            weighted_score = total_score_points / total_credits if score_index != -1 else None

            if weighted_gpa is not None:
                self.gpa_label.setText(f"加权绩点: {weighted_gpa:.5f}")
            else:
                self.gpa_label.setText("加权绩点: 未找到GPA或绩点列")

            if weighted_score is not None:
                self.weighted_score_label.setText(f"加权分数: {weighted_score:.5f}")
            else:
                self.weighted_score_label.setText("加权分数: 未找到总成绩或学分成绩列")
        else:
            self.gpa_label.setText("加权绩点: N/A (没有有效的课程)")
            self.weighted_score_label.setText("加权分数: N/A (没有有效的课程)")
