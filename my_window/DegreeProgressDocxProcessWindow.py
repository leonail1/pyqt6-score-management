"""
这个模块实现了一个Word文档导入工具的图形用户界面。
主要功能包括：
1. 提供一个用户界面来导入和处理Word文档
2. 使用DocxProcess类来处理导入的文档

该模块使用PyQt6来创建图形界面。
"""

from PyQt6.QtCore import Qt
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

    def __init__(self):
        """
        初始化DegreeImportDocxProcessMainWindow对象。

        创建DocxProcess对象并初始化用户界面。
        """
        super().__init__()
        self.docx_processor = DocxProcess(self)
        self.initUI()

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
        """
        self.docx_processor.import_docx()

if __name__ == '__main__':
    """
    程序入口点。

    创建QApplication实例和主窗口，并启动事件循环。
    """
    app = QApplication(sys.argv)
    main_window = DegreeImportDocxProcessMainWindow()
    main_window.show()
    sys.exit(app.exec())