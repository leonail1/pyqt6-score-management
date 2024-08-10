"""
这个模块实现了一个学位进度展示工具的图形用户界面。
主要功能包括：
1. 从JSON文件加载课程信息数据
2. 展示各类课程的学分信息和完成进度
3. 显示课程详细信息的表格

该模块使用PyQt6来创建图形界面。
"""

import json
import os.path
import sys

from PyQt6.QtCore import Qt, QTimer, pyqtSignal
from PyQt6.QtWidgets import (QProgressBar,
                             QLabel, QMessageBox, QHBoxLayout)
from PyQt6.QtWidgets import (QTableWidget, QTableWidgetItem, QHeaderView, QSizePolicy,
                             QDialog, QVBoxLayout, QPushButton, QScrollArea, QWidget,
                             QFrame, QApplication)

from my_window.DegreeImportDocxProcessWindow import DegreeImportDocxProcessMainWindow


class CourseInfoWidget(QWidget):
    """
    显示单个课程类型信息的小部件。
    包括课程类型、必修学分、选修学分和完成进度条。
    """

    def __init__(self, info):
        """
        初始化CourseInfoWidget。

        :param info: 包含课程类型、必修学分和选修学分的元组
        """
        super().__init__()

        layout = QVBoxLayout()

        course_type, required, elective = info

        layout.addWidget(QLabel(f"课程类型: {course_type}", styleSheet="font-weight: bold; font-size: 16px;"))
        layout.addWidget(QLabel(f"必修学分: {required}"))
        layout.addWidget(QLabel(f"选修学分: {elective}"))

        progress = QProgressBar()
        progress.setMaximum(required + elective)
        progress.setValue(required)
        layout.addWidget(progress)

        self.setLayout(layout)


class CourseTableWidget(QTableWidget):
    """
    显示课程详细信息的表格小部件。
    """

    def __init__(self, header, data):
        super().__init__()

        self.setColumnCount(len(header))
        self.setHorizontalHeaderLabels(header)

        self.setRowCount(len(data))
        for row, row_data in enumerate(data):
            has_newline = False
            for col, cell_data in enumerate(row_data):
                item = QTableWidgetItem(str(cell_data))
                item.setTextAlignment(Qt.AlignmentFlag.AlignCenter)
                self.setItem(row, col, item)

                # 检查是否存在换行符
                if '\n' in str(cell_data):
                    has_newline = True

            # 只对存在换行符的行应用自动行高
            if has_newline:
                self.resizeRowToContents(row)

        self.setSortingEnabled(True)
        self.setEditTriggers(QTableWidget.EditTrigger.NoEditTriggers)

        self.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeMode.Interactive)
        self.setMinimumHeight(150)
        self.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Expanding)

        # 调整列宽以适应内容
        self.resizeColumnsToContents()

        # 设置自动换行
        self.setWordWrap(True)


class TableDialog(QDialog):
    """
    显示课程详细信息表格的对话框。
    """

    def __init__(self, parent, table_widget, course_type):
        super().__init__(parent)
        self.setWindowTitle(f"{course_type}详情")
        self.setWindowFlags(self.windowFlags() | Qt.WindowType.WindowStaysOnTopHint)
        self.setAttribute(Qt.WidgetAttribute.WA_DeleteOnClose)

        layout = QVBoxLayout(self)

        self.scroll_area = QScrollArea()
        self.scroll_area.setWidget(table_widget)
        self.scroll_area.setWidgetResizable(True)
        layout.addWidget(self.scroll_area)

        close_button = QPushButton("关闭")
        close_button.clicked.connect(self.close)
        close_button.setDefault(True)
        layout.addWidget(close_button)

        self.setLayout(layout)
        self.resize(800, 600)

        self.table_widget = table_widget

        QTimer.singleShot(0, self.adjust_column_widths)

    def adjust_column_widths(self):
        self.table_widget.resizeColumnsToContents()

        total_width = sum(self.table_widget.columnWidth(i) for i in range(self.table_widget.columnCount()))
        available_width = self.scroll_area.viewport().width() - self.table_widget.verticalScrollBar().width()

        if total_width > available_width:
            self.table_widget.setMinimumWidth(total_width)
        else:
            self.table_widget.setMinimumWidth(available_width)

        self.scroll_area.setVerticalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAsNeeded)
        self.scroll_area.setHorizontalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAsNeeded)
        self.scroll_area.updateGeometry()


class DegreeProgressShowMainWindow(QWidget):
    """
    学位进度展示的主窗口。
    """

    def __init__(self, data):
        super().__init__()
        self.setWindowTitle("课程信息")
        self.resize(800, 600)

        main_layout = QVBoxLayout()

        scroll_area = QScrollArea()
        scroll_content = QWidget()
        scroll_layout = QVBoxLayout(scroll_content)

        self.table_widgets = []

        for i, item in enumerate(data):
            frame = QFrame()
            frame.setFrameShape(QFrame.Shape.StyledPanel)
            frame.setFrameShadow(QFrame.Shadow.Raised)
            frame_layout = QVBoxLayout(frame)

            info_widget = CourseInfoWidget(item['info'])
            frame_layout.addWidget(info_widget)

            show_details_button = QPushButton(f"显示{item['info'][0]}详情")
            frame_layout.addWidget(show_details_button)

            table_widget = CourseTableWidget(item['table']['header'], item['table']['data'])
            self.table_widgets.append((table_widget, item['info'][0]))

            show_details_button.clicked.connect(
                lambda checked, w=table_widget, t=item['info'][0]: self.show_table_dialog(w, t))

            scroll_layout.addWidget(frame)

            if i < len(data) - 1:
                line = QFrame()
                line.setFrameShape(QFrame.Shape.HLine)
                line.setFrameShadow(QFrame.Shadow.Sunken)
                scroll_layout.addWidget(line)

        scroll_content.setLayout(scroll_layout)
        scroll_area.setWidget(scroll_content)
        scroll_area.setWidgetResizable(True)

        main_layout.addWidget(scroll_area)

        # 添加关闭按钮
        close_button = QPushButton("关闭窗口")
        close_button.clicked.connect(self.close)
        close_button.setDefault(True)  # 设置为默认按钮

        # 创建按钮布局
        button_layout = QHBoxLayout()
        button_layout.addStretch()  # 添加弹性空间，使按钮靠右对齐
        button_layout.addWidget(close_button)

        main_layout.addLayout(button_layout)

        self.setLayout(main_layout)

    def show_table_dialog(self, table_widget, course_type):
        dialog = TableDialog(self, table_widget, course_type)
        dialog.setModal(True)
        dialog.show()


class DataManager:
    """
    管理数据加载和错误处理的类。
    """

    def __init__(self):
        """
        初始化DataManager。
        """
        self.data = None
        self.file_path = None

    def set_file_path(self, file_path):
        """
        设置数据文件的路径。

        :param file_path: 数据文件的路径
        """
        self.file_path = file_path

    def load_data(self):
        """
        从文件加载数据。

        :return: 如果加载成功返回True，否则返回False
        """
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
        """
        显示错误对话框。

        :param error_type: 错误类型
        :param error_message: 错误消息
        """
        error_dialog = QMessageBox()
        error_dialog.setIcon(QMessageBox.Icon.Critical)
        error_dialog.setText(error_type)
        error_dialog.setInformativeText(error_message)
        error_dialog.setWindowTitle("数据加载错误")
        error_dialog.exec()

    def get_data(self):
        """
        获取加载的数据。

        :return: 加载的数据
        """
        return self.data


class DegreeProgressWidget(QWidget):
    import_finished = pyqtSignal()

    def __init__(self, parent=None):
        super().__init__(parent)
        self.data_manager = DataManager()
        self.setup_data_manager()
        self.progress_window = None
        self.import_window = None

    def setup_data_manager(self):
        current_dir = os.path.dirname(os.path.abspath(__file__))
        data_file_path = os.path.join(current_dir, "..", "config", "degree_progress.json")
        self.data_manager.set_file_path(data_file_path)

    def run_import_program(self):
        self.import_window = DegreeImportDocxProcessMainWindow()
        self.import_window.show()
        self.import_window.import_finished.connect(self.on_import_finished)

    def on_import_finished(self):
        if self.import_window:
            self.import_window.close()
        self.import_finished.emit()

    def show_degree_progress(self):
        if self.data_manager.load_data():
            self.progress_window = DegreeProgressShowMainWindow(self.data_manager.get_data())
            self.progress_window.show()
            return self.progress_window
        else:
            QMessageBox.critical(None, "错误", "无法加载数据。")
            return None

    def start(self):
        if not self.data_manager.load_data():
            msg_box = QMessageBox()
            msg_box.setIcon(QMessageBox.Icon.Question)
            msg_box.setText("数据加载失败。是否运行文件导入程序？")
            msg_box.setWindowTitle("数据加载失败")
            msg_box.setStandardButtons(QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No)
            msg_box.setDefaultButton(QMessageBox.StandardButton.Yes)

            if msg_box.exec() == QMessageBox.StandardButton.Yes:
                self.run_import_program()
                self.import_finished.connect(self.show_degree_progress)
                return None  # 返回None，因为窗口还没有准备好
            else:
                return None
        else:
            return self.show_degree_progress()


def create_degree_progress_window(parent=None):
    """
    创建并返回学位进度窗口，可以从其他地方调用而不会导致事件循环冲突。

    :param parent: 父窗口，默认为None
    :return: DegreeProgressShowMainWindow实例或None
    """
    widget = DegreeProgressWidget(parent)
    return widget.start()


def main():
    app = QApplication(sys.argv)
    window = create_degree_progress_window()
    if window:
        window.show()
        sys.exit(app.exec())
    else:
        sys.exit(0)


if __name__ == "__main__":
    main()
