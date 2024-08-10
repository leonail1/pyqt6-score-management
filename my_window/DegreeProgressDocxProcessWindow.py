from PyQt6.QtCore import Qt
from PyQt6.QtWidgets import QMainWindow, QApplication, QHBoxLayout, QLabel, QPushButton, QMessageBox, QVBoxLayout, \
    QWidget
import sys
import os

sys.path.append(os.getcwd())

from degree_process.docx_process import DocxProcess


class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.docx_processor = DocxProcess(self)
        self.initUI()

    def initUI(self):
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

        # 导入按钮
        self.import_button = QPushButton("导入文件")
        self.import_button.clicked.connect(self.handle_import)
        button_layout.addWidget(self.import_button)

        # 取消按钮
        cancel_button = QPushButton("取消")
        cancel_button.clicked.connect(self.close)
        button_layout.addWidget(cancel_button)

        layout.addLayout(button_layout)

        central_widget.setLayout(layout)

    def handle_import(self):
        self.docx_processor.import_docx()

if __name__ == '__main__':
    app = QApplication(sys.argv)
    main_window = MainWindow()
    main_window.show()
    sys.exit(app.exec())