"""
这个模块实现了一个Word文档导入工具的图形用户界面。
主要功能包括：
1. 提供一个用户界面来导入和处理Word文档
2. 使用DocxProcess类来处理导入的文档

该模块使用PyQt6来创建图形界面。
"""

from PyQt6.QtCore import Qt, pyqtSignal
from PyQt6.QtWidgets import QMainWindow, QApplication, QHBoxLayout, QLabel, QPushButton, QMessageBox, QVBoxLayout, \
    QWidget
import sys
import os

# 将当前工作目录添加到系统路径
sys.path.append(os.getcwd())

from degree_process.docx_process import DocxProcess


class DegreeImportDocxProcessMainWindow(QMainWindow):
    """
    Word文档导入工具的主窗口类。

    这个类创建了一个图形用户界面，允许用户导入和处理Word文档。
    """

    import_finished = pyqtSignal()  # 添加导入完成信号

    def __init__(self, student_id):
        """
        初始化DegreeImportDocxProcessMainWindow对象。

        创建DocxProcess对象并初始化用户界面。
        """
        super().__init__()
        # print(f"DegreeImportDocxProcessMainWindow.__init__ 被调用，学生ID: {student_id}")
        self.student_id = student_id
        self.docx_processor = DocxProcess(self)
        self.initUI()
        # print("DegreeImportDocxProcessMainWindow 初始化完成")
        # print(f"窗口大小: {self.size()}")
        # print(f"窗口位置: {self.pos()}")

    def initUI(self):
        """
        初始化用户界面。

        设置窗口标题、大小，并创建各种UI元素，如标签、按钮等。
        """
        self.setWindowTitle('Word文档导入工具')
        self.setGeometry(300, 300, 400, 200)

        central_widget = QWidget()
        self.setCentralWidget(central_widget)

        layout = QVBoxLayout()

        # 添加提示标签
        info_label = QLabel("请选择要导入的 .docx 文件：")
        info_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        layout.addWidget(info_label)

        # 添加文件名显示标签
        self.file_label = QLabel("未选择文件")
        self.file_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        layout.addWidget(self.file_label)

        button_layout = QHBoxLayout()

        # 添加导入按钮
        self.import_button = QPushButton("导入文件")
        self.import_button.clicked.connect(self.handle_import)
        button_layout.addWidget(self.import_button)

        # 添加取消按钮
        cancel_button = QPushButton("取消")
        cancel_button.clicked.connect(self.close)
        button_layout.addWidget(cancel_button)

        layout.addLayout(button_layout)

        central_widget.setLayout(layout)

    def handle_import(self):
        """
        处理导入按钮点击事件。

        调用DocxProcess对象的import_docx方法来导入和处理Word文档。
        如果导入成功，发出import_finished信号并关闭窗口。
        如果导入失败，显示错误消息。
        """
        result = self.docx_processor.import_docx(student_id=self.student_id)
        if result is not None:
            self.import_finished.emit()
            self.close()
        else:
            QMessageBox.critical(self, "导入失败", "文档导入失败，请检查文件格式或重试。")

    def closeEvent(self, event):
        self.import_finished.emit()
        event.accept()


if __name__ == '__main__':
    """
    程序入口点。

    创建QApplication实例和主窗口，并启动事件循环。
    """
    app = QApplication(sys.argv)
    main_window = DegreeImportDocxProcessMainWindow(student_id="37220222203691")
    main_window.show()
    sys.exit(app.exec())
