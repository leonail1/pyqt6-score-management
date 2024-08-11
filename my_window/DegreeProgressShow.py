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


# 导入 DegreeImportDocxProcessMainWindow 类，该类应该定义了文件导入程序的逻辑


class CourseInfoWidget(QWidget):
    """
    显示单个课程类型信息的小部件。
    包括课程类型、必修学分、选修学分和完成进度条。
    """

    def __init__(self, info):
        """
        初始化CourseInfoWidget。

        :param info: 包含课程类型、必修学分和选修学分的元组，例如：("公共基础课", 10, 6)
        """
        super().__init__()

        layout = QVBoxLayout()  # 使用垂直布局

        course_type, required, elective = info  # 解包课程信息

        # 添加课程类型标签，设置字体样式
        layout.addWidget(QLabel(f"课程类型: {course_type}", styleSheet="font-weight: bold; font-size: 16px;"))
        # 添加必修学分标签
        layout.addWidget(QLabel(f"必修学分: {required}"))
        # 添加选修学分标签
        layout.addWidget(QLabel(f"选修学分: {elective}"))

        progress = QProgressBar()  # 创建进度条
        progress.setMaximum(required + elective)  # 设置进度条最大值
        progress.setValue(required)  # 设置进度条当前值
        layout.addWidget(progress)  # 添加进度条到布局

        self.setLayout(layout)  # 将布局应用到小部件


class CourseTableWidget(QTableWidget):
    """
    显示课程详细信息的表格小部件。
    """

    def __init__(self, header, data):
        """
        初始化 CourseTableWidget。

        :param header: 表格头部，列表形式，例如：["课程代码", "课程名称", "学分", "成绩"]
        :param data: 表格数据，二维列表形式，例如：
                     [
                         ["00001", "高等数学", 5, "90"],
                         ["00002", "大学英语", 4, "85"]
                     ]
        """
        super().__init__()

        self.setColumnCount(len(header))  # 设置列数
        self.setHorizontalHeaderLabels(header)  # 设置表头

        self.setRowCount(len(data))  # 设置行数
        for row, row_data in enumerate(data):  # 遍历数据，填充表格
            has_newline = False  # 用于标记该行是否存在换行符
            for col, cell_data in enumerate(row_data):
                item = QTableWidgetItem(str(cell_data))  # 创建表格项
                item.setTextAlignment(Qt.AlignmentFlag.AlignCenter)  # 设置文本居中
                self.setItem(row, col, item)  # 将表格项添加到表格

                # 检查是否存在换行符
                if '\n' in str(cell_data):
                    has_newline = True

            # 只对存在换行符的行应用自动行高
            if has_newline:
                self.resizeRowToContents(row)

        self.setSortingEnabled(True)  # 启用排序功能
        self.setEditTriggers(QTableWidget.EditTrigger.NoEditTriggers)  # 禁止编辑表格

        self.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeMode.Interactive)  # 设置表头可交互调整大小
        self.setMinimumHeight(150)  # 设置最小高度
        self.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Expanding)  # 设置大小策略，使其可以扩展

        # 调整列宽以适应内容
        self.resizeColumnsToContents()

        # 设置自动换行
        self.setWordWrap(True)


class TableDialog(QDialog):
    """
    显示课程详细信息表格的对话框。
    """

    def __init__(self, parent, table_widget, course_type):
        """
        初始化 TableDialog。

        :param parent: 父窗口
        :param table_widget: 要显示的 CourseTableWidget 实例
        :param course_type: 课程类型，用于设置窗口标题
        """
        super().__init__(parent)
        self.setWindowTitle(f"{course_type}详情")  # 设置窗口标题
        self.setWindowFlags(self.windowFlags() | Qt.WindowType.WindowStaysOnTopHint)  # 设置窗口始终置顶
        # 移除 WA_DeleteOnClose 属性，避免关闭对话框时自动删除 table_widget
        # self.setAttribute(Qt.WidgetAttribute.WA_DeleteOnClose)

        layout = QVBoxLayout(self)  # 使用垂直布局

        self.scroll_area = QScrollArea()  # 创建滚动区域
        self.scroll_area.setWidget(table_widget)  # 将表格添加到滚动区域
        self.scroll_area.setWidgetResizable(True)  # 设置滚动区域内容可调整大小
        layout.addWidget(self.scroll_area)  # 将滚动区域添加到布局

        close_button = QPushButton("关闭")  # 创建关闭按钮
        close_button.clicked.connect(self.close)  # 连接关闭按钮的点击信号到关闭对话框
        close_button.setDefault(True)  # 设置为默认按钮
        layout.addWidget(close_button)  # 将关闭按钮添加到布局

        self.setLayout(layout)  # 将布局应用到对话框
        self.resize(800, 600)  # 设置对话框大小

        self.table_widget = table_widget  # 保存 table_widget 实例

        # 使用 QTimer.singleShot 在事件循环的下一个周期调整列宽，
        # 确保在对话框完全显示后才进行调整
        QTimer.singleShot(0, self.adjust_column_widths)

    def adjust_column_widths(self):
        """
        调整表格列宽以适应内容，并设置滚动条策略。
        """
        self.table_widget.resizeColumnsToContents()  # 调整列宽以适应内容

        # 计算表格总宽度和可用宽度
        total_width = sum(self.table_widget.columnWidth(i) for i in range(self.table_widget.columnCount()))
        available_width = self.scroll_area.viewport().width() - self.table_widget.verticalScrollBar().width()

        # 如果表格总宽度大于可用宽度，则设置表格最小宽度为总宽度，否则设置为可用宽度
        if total_width > available_width:
            self.table_widget.setMinimumWidth(total_width)
        else:
            self.table_widget.setMinimumWidth(available_width)

        # 设置滚动条策略
        self.scroll_area.setVerticalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAsNeeded)
        self.scroll_area.setHorizontalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAsNeeded)
        self.scroll_area.updateGeometry()  # 更新滚动区域几何形状


class DegreeProgressShowMainWindow(QWidget):
    """
    学位进度展示的主窗口。
    """

    def __init__(self, data):
        """
        初始化 DegreeProgressShowMainWindow。

        :param data: 课程数据，列表形式，每个元素是一个字典，包含以下键值：
                     - 'info': 课程类型信息，元组形式，例如：("公共基础课", 10, 6)
                     - 'table': 课程详细信息，字典形式，包含以下键值：
                       - 'header': 表格头部，列表形式
                       - 'data': 表格数据，二维列表形式
        """
        super().__init__()
        self.setWindowTitle("课程信息")  # 设置窗口标题
        self.resize(800, 600)  # 设置窗口大小

        main_layout = QVBoxLayout()  # 使用垂直布局

        scroll_area = QScrollArea()  # 创建滚动区域
        scroll_content = QWidget()  # 创建滚动区域内容部件
        scroll_layout = QVBoxLayout(scroll_content)  # 使用垂直布局

        self.table_dialogs = {}  # 使用字典存储对话框，避免重复创建

        for i, item in enumerate(data):  # 遍历课程数据
            frame = QFrame()  # 创建框架，用于包含每个课程类型的信息
            frame.setFrameShape(QFrame.Shape.StyledPanel)  # 设置框架样式
            frame.setFrameShadow(QFrame.Shadow.Raised)  # 设置框架阴影
            frame_layout = QVBoxLayout(frame)  # 使用垂直布局

            info_widget = CourseInfoWidget(item['info'])  # 创建课程信息小部件
            frame_layout.addWidget(info_widget)  # 将课程信息小部件添加到框架布局

            show_details_button = QPushButton(f"显示{item['info'][0]}详情")  # 创建显示详情按钮
            frame_layout.addWidget(show_details_button)  # 将按钮添加到框架布局

            table_widget = CourseTableWidget(item['table']['header'], item['table']['data'])  # 创建课程表格
            # 创建对话框并存储
            table_dialog = TableDialog(self, table_widget, item['info'][0])
            self.table_dialogs[item['info'][0]] = table_dialog

            # 连接按钮的点击信号到显示对话框的槽函数
            show_details_button.clicked.connect(
                lambda checked, course_type=item['info'][0]: self.show_table_dialog(course_type)
            )

            scroll_layout.addWidget(frame)  # 将框架添加到滚动区域布局

            # 添加分隔线
            if i < len(data) - 1:
                line = QFrame()
                line.setFrameShape(QFrame.Shape.HLine)
                line.setFrameShadow(QFrame.Shadow.Sunken)
                scroll_layout.addWidget(line)

        scroll_content.setLayout(scroll_layout)  # 将布局应用到滚动区域内容部件
        scroll_area.setWidget(scroll_content)  # 将内容部件添加到滚动区域
        scroll_area.setWidgetResizable(True)  # 设置滚动区域内容可调整大小

        main_layout.addWidget(scroll_area)  # 将滚动区域添加到主布局

        # 添加关闭按钮
        close_button = QPushButton("关闭窗口")
        close_button.clicked.connect(self.close)  # 连接关闭按钮的点击信号到关闭窗口
        close_button.setDefault(True)  # 设置为默认按钮

        # 创建按钮布局
        button_layout = QHBoxLayout()
        button_layout.addStretch()  # 添加弹性空间，使按钮靠右对齐
        button_layout.addWidget(close_button)  # 将关闭按钮添加到按钮布局

        main_layout.addLayout(button_layout)  # 将按钮布局添加到主布局

        self.setLayout(main_layout)  # 将主布局应用到窗口

    def show_table_dialog(self, course_type):
        """
        显示对应课程类型的对话框。

        :param course_type: 课程类型
        """
        dialog = self.table_dialogs[course_type]  # 获取对应的对话框
        dialog.show()  # 显示对话框


class DataManager:
    """
    管理数据加载和错误处理的类。
    """

    def __init__(self):
        """
        初始化 DataManager。
        """
        self.data = None  # 数据存储
        self.file_path = None  # 数据文件路径

    def set_file_path(self, file_path):
        """
        设置数据文件的路径。

        :param file_path: 数据文件的路径
        """
        self.file_path = file_path

    def load_data(self):
        """
        从文件加载数据。

        :return: 如果加载成功返回 True，否则返回 False
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
    """
    管理学位进度窗口的部件，包括数据加载、文件导入和窗口显示。
    """
    import_finished = pyqtSignal()  # 定义导入完成信号

    def __init__(self, student_id, parent=None):
        """
        初始化 DegreeProgressWidget。

        :param student_id: 学生ID
        :param parent: 父窗口
        """
        super().__init__(parent)
        self.data_manager = DataManager()  # 创建 DataManager 实例
        self.setup_data_manager(student_id=student_id)  # 设置数据文件路径
        self.progress_window = None  # 学位进度窗口
        self.import_window = None  # 文件导入窗口

    def setup_data_manager(self, student_id):
        """
        根据学生ID设置数据文件路径。

        :param student_id: 学生ID
        """
        current_dir = os.path.dirname(os.path.abspath(__file__))
        data_file_path = os.path.join(current_dir, "..", "config", "degree_progress_" + str(student_id) + ".json")
        self.data_manager.set_file_path(data_file_path)

    def run_import_program(self, student_id):
        """
        运行文件导入程序。

        :param student_id: 学生ID
        """
        self.import_window = DegreeImportDocxProcessMainWindow(student_id=student_id)
        self.import_window.show()
        self.import_window.import_finished.connect(self.on_import_finished)  # 连接导入完成信号到处理函数

    def on_import_finished(self):
        """
        文件导入完成后的处理函数。
        """
        if self.import_window:
            self.import_window.close()  # 关闭导入窗口
        self.import_finished.emit()  # 发射导入完成信号

    def show_degree_progress(self):
        """
        显示学位进度窗口。

        :return: DegreeProgressShowMainWindow 实例或 None
        """
        if self.data_manager.load_data():
            self.progress_window = DegreeProgressShowMainWindow(self.data_manager.get_data())
            self.progress_window.show()
            return self.progress_window
        else:
            QMessageBox.critical(None, "错误", "无法加载数据。")
            return None

    def start(self, student_id):
        """
        启动程序，尝试加载数据，如果加载失败则询问是否运行文件导入程序。

        :param student_id: 学生ID
        :return: DegreeProgressShowMainWindow 实例或 None
        """
        if not self.data_manager.load_data():
            msg_box = QMessageBox()
            msg_box.setIcon(QMessageBox.Icon.Question)
            msg_box.setText("数据加载失败。是否运行文件导入程序？")
            msg_box.setWindowTitle("数据加载失败")
            msg_box.setStandardButtons(QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No)
            msg_box.setDefaultButton(QMessageBox.StandardButton.Yes)

            if msg_box.exec() == QMessageBox.StandardButton.Yes:
                self.run_import_program(student_id=student_id)
                self.import_finished.connect(self.show_degree_progress)  # 连接导入完成信号到显示窗口
                return None  # 返回 None，因为窗口还没有准备好
            else:
                return None
        else:
            return self.show_degree_progress()


def create_degree_progress_window(student_id, parent=None):
    """
    创建并返回学位进度窗口，可以从其他地方调用而不会导致事件循环冲突。

    :param student_id: 学分进度对应的学号
    :param parent: 父窗口，默认为None
    :return: DegreeProgressShowMainWindow实例或None
    """
    widget = DegreeProgressWidget(parent=parent, student_id=student_id)
    return widget.start(student_id=student_id)


def __main():
    app = QApplication(sys.argv)
    window = create_degree_progress_window(student_id="37220222203691")
    if window:
        window.show()
        sys.exit(app.exec())
    else:
        sys.exit(0)


if __name__ == "__main__":
    __main()
