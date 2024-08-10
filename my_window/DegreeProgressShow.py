import json
import os.path
import sys
from PyQt6.QtWidgets import (QApplication, QWidget, QVBoxLayout, QHBoxLayout,
                             QTableWidget, QTableWidgetItem, QProgressBar,
                             QLabel, QScrollArea, QSizePolicy, QPushButton, QFrame, QHeaderView, QMessageBox)
from PyQt6.QtCore import Qt
from PyQt6.QtGui import QPalette, QColor
from degree_process.docx_process import DocxProcess


class CourseInfoWidget(QWidget):
    def __init__(self, info):
        super().__init__()
        layout = QVBoxLayout()

        course_type, required, elective = info
        layout.addWidget(QLabel(f"课程类型: {course_type}", styleSheet="font-weight: bold; font-size: 16px;"))
        layout.addWidget(QLabel(f"必修学分: {required}"))
        layout.addWidget(QLabel(f"选修学分: {elective}"))

        progress = QProgressBar()
        progress.setMaximum(required + elective)
        progress.setValue(required)  # 假设当前完成的是必修学分
        layout.addWidget(progress)

        self.setLayout(layout)


class CourseTableWidget(QTableWidget):
    def __init__(self, header, data):
        super().__init__()
        self.setColumnCount(len(header))
        self.setHorizontalHeaderLabels(header)
        self.setRowCount(len(data))

        for row, row_data in enumerate(data):
            for col, cell_data in enumerate(row_data):
                item = QTableWidgetItem(str(cell_data))
                item.setTextAlignment(Qt.AlignmentFlag.AlignCenter)  # 居中显示
                self.setItem(row, col, item)

        self.setSortingEnabled(True)
        self.setEditTriggers(QTableWidget.EditTrigger.NoEditTriggers)  # 设置表格为只读

        # 调整表格大小以适应所有内容
        self.resizeColumnsToContents()
        self.resizeRowsToContents()

        # 设置最小列宽和行高
        for col in range(self.columnCount()):
            self.setColumnWidth(col, max(30, self.columnWidth(col)))
        for row in range(self.rowCount()):
            self.setRowHeight(row, max(15, self.rowHeight(row)))

        # 设置表头样式
        self.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeMode.Stretch)

        # 设置表格的最小高度
        self.setMinimumHeight(150)  # 设置最小高度为150像素

        self.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Expanding)


class MainWindow(QWidget):
    def __init__(self, data):
        super().__init__()
        self.setWindowTitle("课程信息")
        self.resize(800, 600)

        main_layout = QVBoxLayout()

        # 创建一个滚动区域来容纳所有内容
        scroll_area = QScrollArea()
        scroll_content = QWidget()
        scroll_layout = QVBoxLayout(scroll_content)

        self.table_widgets = []  # 存储所有表格部件

        for i, item in enumerate(data):
            # 创建一个框架来包含每个课程类型的所有内容
            frame = QFrame()
            frame.setFrameShape(QFrame.Shape.StyledPanel)
            frame.setFrameShadow(QFrame.Shadow.Raised)
            frame_layout = QVBoxLayout(frame)

            # 创建课程信息部件
            info_widget = CourseInfoWidget(item['info'])
            frame_layout.addWidget(info_widget)

            # 创建"显示详情"按钮
            show_details_button = QPushButton(f"显示{item['info'][0]}详情")
            frame_layout.addWidget(show_details_button)

            # 创建表格部件（初始隐藏）
            table_widget = CourseTableWidget(item['table']['header'], item['table']['data'])
            table_widget.hide()
            frame_layout.addWidget(table_widget)
            self.table_widgets.append(table_widget)

            # 连接按钮点击事件
            show_details_button.clicked.connect(
                lambda checked, w=table_widget, b=show_details_button, t=item['info'][0]: self.toggle_table(w, b, t))

            scroll_layout.addWidget(frame)

            # 在最后一个项目之前添加分隔线
            if i < len(data) - 1:
                line = QFrame()
                line.setFrameShape(QFrame.Shape.HLine)
                line.setFrameShadow(QFrame.Shadow.Sunken)
                scroll_layout.addWidget(line)

        scroll_content.setLayout(scroll_layout)
        scroll_area.setWidget(scroll_content)
        scroll_area.setWidgetResizable(True)

        main_layout.addWidget(scroll_area)
        self.setLayout(main_layout)

    def toggle_table(self, table_widget, button, course_type):
        if table_widget.isVisible():
            table_widget.hide()
            button.setText(f"显示{course_type}详情")
        else:
            table_widget.show()
            button.setText(f"隐藏{course_type}详情")


class DataManager:
    def __init__(self):
        self.data = None
        self.file_path = None

    def set_file_path(self, file_path):
        self.file_path = file_path

    def load_data(self):
        if not self.file_path:
            self.show_error_dialog("路径错误", "数据文件路径未设置。")
            return False

        try:
            with open(self.file_path, 'r', encoding='utf-8') as file:
                self.data = json.load(file)
            return True
        except FileNotFoundError:
            self.show_error_dialog("文件不存在", f"无法找到数据文件：\n{self.file_path}")
        except json.JSONDecodeError:
            self.show_error_dialog("JSON 格式错误", f"数据文件格式不正确：\n{self.file_path}")
        except Exception as e:
            self.show_error_dialog("未知错误", f"加载数据时发生未知错误：\n{str(e)}")
        return False

    @staticmethod
    def show_error_dialog(error_type, error_message):
        error_dialog = QMessageBox()
        error_dialog.setIcon(QMessageBox.Icon.Critical)
        error_dialog.setText(error_type)
        error_dialog.setInformativeText(error_message)
        error_dialog.setWindowTitle("数据加载错误")
        error_dialog.exec()

    def get_data(self):
        return self.data

if __name__ == '__main__':
    app = QApplication(sys.argv)

    # 创建 DataManager 实例
    data_manager = DataManager()

    # 设置数据文件路径
    current_dir = os.path.dirname(os.path.abspath(__file__))
    data_file_path = os.path.join(current_dir, "..", "config", "degree_progress.json")
    data_manager.set_file_path(data_file_path)

    # 尝试加载数据
    if data_manager.load_data():
        # 数据加载成功，创建并显示主窗口
        window = MainWindow(data_manager.get_data())
        window.show()
        sys.exit(app.exec())
    else:
        sys.exit(1)