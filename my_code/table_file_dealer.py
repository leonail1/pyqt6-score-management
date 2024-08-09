import json
import os
import subprocess
import sys
import traceback

import pandas as pd
from PyQt6.QtWidgets import QFileDialog, QMessageBox, QInputDialog, QLineEdit, QDialog, QVBoxLayout, QLabel, \
    QPushButton, QHBoxLayout

from my_window.StudentInfoWindow import StudentInfoWindow
from scraper.scraper import main as run_grade_scraper


class FileDealer:
    def __init__(self, parent):
        self.file_from_scraper = False
        self.student_score_analyzer = None
        self.parent = parent

    def set_default_student_id(self, student_id_input: QLineEdit):
        """
        从配置文件中读取上一次输入的学号信息
        """
        data_dir = os.path.abspath(os.path.join(os.path.dirname(__file__), '..', 'config'))
        config_file = os.path.join(data_dir, "user_config.json")
        default_id = ""

        if os.path.exists(config_file):
            try:
                with open(config_file, 'r') as f:
                    config = json.load(f)
                    stored_id = config.get("student_id_input", "")
                    if len(stored_id) == 14 and stored_id.isdigit():
                        default_id = stored_id
                    else:
                        # 重置配置文件中的 student_id_input
                        config["student_id_input"] = ""
                        with open(config_file, 'w') as f:
                            json.dump(config, f, indent=4)
            except json.JSONDecodeError:
                print("配置文件格式错误")
            except IOError:
                print("读取配置文件时发生错误")

        student_id_input.setText(default_id)

    def process_student_id(self, student_id_input):
        """
        输入学号后连接到的操作，读取或导入文件
        """
        student_id = student_id_input.text()
        if len(student_id) != 14:
            QMessageBox.warning(self.parent, "错误", "学号必须是14位数字")
            return

        # 保存当前输入的学号到配置文件
        self.save_student_id_to_config(student_id)

        data_dir = os.path.abspath(os.path.join(os.path.dirname(__file__), '..', 'data'))
        file_path = os.path.join(data_dir, f"{student_id}.json")

        if os.path.exists(file_path):
            msg_box = QMessageBox()
            msg_box.setWindowTitle("学生数据已存在")
            msg_box.setText(f"找到学生数据: {student_id}\n请选择操作：")
            msg_box.setStandardButtons(QMessageBox.StandardButton.Open |
                                       QMessageBox.StandardButton.Save |
                                       QMessageBox.StandardButton.Discard |
                                       QMessageBox.StandardButton.Cancel)
            msg_box.button(QMessageBox.StandardButton.Save).setText("Overwrite")
            msg_box.button(QMessageBox.StandardButton.Discard).setText("Delete")

            reply = msg_box.exec()

            if reply == QMessageBox.StandardButton.Open:
                self.load_and_display_student_data(student_id)
            elif reply == QMessageBox.StandardButton.Save:
                if not self.import_file(student_id_input=student_id):
                    return  # 如果导入被取消，直接返回
                QMessageBox.information(self.parent, "成功", "数据已成功覆盖")
            elif reply == QMessageBox.StandardButton.Discard:
                confirm = QMessageBox.question(self.parent, "确认删除",
                                               "确定要删除该学生数据吗？此操作不可撤销。",
                                               QMessageBox.StandardButton.Yes |
                                               QMessageBox.StandardButton.No)
                if confirm == QMessageBox.StandardButton.Yes:
                    os.remove(file_path)
                    QMessageBox.information(self.parent, "成功", "学生数据已删除")
            else:  # Cancel
                return
        else:
            reply = QMessageBox.question(self.parent, "学生数据不存在",
                                         f"未找到学生 {student_id} 的数据。是否要导入新数据？",
                                         QMessageBox.StandardButton.Yes |
                                         QMessageBox.StandardButton.No)
            if reply == QMessageBox.StandardButton.Yes:
                if not self.import_file(student_id_input=student_id):
                    return  # 如果导入被取消，直接返回
                QMessageBox.information(self.parent, "成功", "新学生数据已导入")
                self.load_and_display_student_data(student_id)

    def load_and_display_student_data(self, student_id):
        # QMessageBox.information(self.parent, "加载数据", f"已加载学生 {student_id} 的数据")

        # 创建新窗口
        self.student_info_window = StudentInfoWindow("setup_score_list_view_ui", student_id=student_id)

        # 设置新窗口为模态：在新窗口退出前不可编辑主窗口（可选）
        self.student_info_window.setModal(False)

        # 显示新窗口
        self.student_info_window.show()

        # 如果你希望新窗口在关闭前阻塞主程序，可以使用 exec() 而不是 show()
        # self.student_info_window.exec()

    def save_student_id_to_config(self, student_id):
        """
        将新输入的学号保存到user_config.json
        """
        data_dir = os.path.abspath(os.path.join(os.path.dirname(__file__), '..', 'config'))
        config_file = os.path.join(data_dir, "user_config.json")
        config = {}

        if os.path.exists(config_file):
            try:
                with open(config_file, 'r') as f:
                    config = json.load(f)
            except json.JSONDecodeError:
                print("配置文件格式错误，将创建新的配置")
            except IOError:
                print("读取配置文件时发生错误，将创建新的配置")

        config["student_id_input"] = student_id

        try:
            os.makedirs(os.path.dirname(config_file), exist_ok=True)
            with open(config_file, 'w') as f:
                json.dump(config, f, indent=4)
        except IOError:
            print("保存配置文件时发生错误")

    def input_student_info(self, name_input: str = None, student_id_input: str = None):
        """
        获取姓名、学号输入
        """
        name, student_id = name_input, student_id_input
        if student_id and name:
            return name, student_id
        elif student_id:
            name, ok1 = QInputDialog.getText(self.parent, "输入姓名", "请输入你的姓名：", QLineEdit.EchoMode.Normal)
            if ok1:
                return name, student_id
            else:
                QMessageBox.warning(self.parent, "Warning", "未输入姓名，操作已取消。")
        elif name:
            student_id, ok2 = QInputDialog.getText(self.parent, "输入学号", "请输入你的学号：",
                                                   QLineEdit.EchoMode.Normal)
            if ok2:
                return name, student_id
            else:
                QMessageBox.warning(self.parent, "Warning", "未输入学号，操作已取消。")
        else:
            name, ok1 = QInputDialog.getText(self.parent, "输入姓名", "请输入你的姓名：", QLineEdit.EchoMode.Normal)
            if ok1:
                student_id, ok2 = QInputDialog.getText(self.parent, "输入学号", "请输入你的学号：",
                                                       QLineEdit.EchoMode.Normal)
                if ok2:
                    return name, student_id
                else:
                    QMessageBox.warning(self.parent, "Warning", "未输入学号，操作已取消。")
            else:
                QMessageBox.warning(self.parent, "Warning", "未输入姓名，操作已取消。")

        return None, None

    def import_file(self, student_id_input: str = None, name_input: str = None):
        """
        导入教务成绩文件
        """
        name, student_id = self.input_student_info(student_id_input=student_id_input, name_input=name_input)
        if not name or not student_id:
            return

        # 检查是否存在同名文件
        data_dir = os.path.abspath(os.path.join(os.path.dirname(__file__), '..', 'data'))
        json_file_name = os.path.join(data_dir, f"{student_id}.json")

        if os.path.exists(json_file_name):
            reply = QMessageBox.question(self.parent, '文件已存在',
                                         f"学号 {student_id} 的文件已存在。是否覆盖？",
                                         QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
                                         QMessageBox.StandardButton.No)
            if reply == QMessageBox.StandardButton.No:
                QMessageBox.information(self.parent, "操作取消", "导入操作已取消。")
                return

        # 询问用户是否使用爬虫导入
        scraper_reply = QMessageBox.question(self.parent, '选择导入方式',
                                             "是否使用爬虫导入数据？",
                                             QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
                                             QMessageBox.StandardButton.No)

        if scraper_reply == QMessageBox.StandardButton.Yes:
            self.file_from_scraper = True

            # 获取当前脚本的目录
            current_dir = os.path.dirname(os.path.abspath(__file__))
            # 构建爬虫脚本的路径
            scraper_path = os.path.join(current_dir, '..', 'scraper', 'scraper.py')

            # 创建一个自定义对话框
            dialog = QDialog(self.parent)
            dialog.setWindowTitle("爬虫脚本位置")
            layout = QVBoxLayout()

            # 添加说明文本
            label = QLabel(f"爬虫脚本位置：\n{scraper_path}\n\n请手动运行该脚本，完成后点击下方的确认按钮。")
            layout.addWidget(label)

            # 创建一个水平布局来放置按钮
            button_layout = QHBoxLayout()

            # 添加确认按钮
            confirm_button = QPushButton("确认已运行爬虫")
            confirm_button.clicked.connect(dialog.accept)
            button_layout.addWidget(confirm_button)

            # 添加取消按钮
            cancel_button = QPushButton("取消")
            cancel_button.clicked.connect(dialog.reject)
            button_layout.addWidget(cancel_button)

            # 将按钮布局添加到主布局
            layout.addLayout(button_layout)

            dialog.setLayout(layout)

            # 显示对话框
            result = dialog.exec()

            if result != QDialog.DialogCode.Accepted:
                # 用户取消操作
                self.file_from_scraper = False
                QMessageBox.information(self.parent, "操作取消", "您已取消运行爬虫操作。")
                return False  # 返回 False 表示操作被取消

        if self.file_from_scraper:
            # 读取爬虫生成的文件
            scraper_file = os.path.abspath(
                os.path.join(os.path.dirname(__file__), '..', 'scraper', 'table_contents', 'all_tables_content.xlsx'))
            try:
                df = pd.read_excel(scraper_file, sheet_name='总表')
            except Exception as e:
                QMessageBox.critical(self.parent, "Error", f"无法读取爬虫生成的文件: {str(e)}")
                return
        else:
            # 原有的文件选择逻辑
            file_name, _ = QFileDialog.getOpenFileName(
                self.parent,
                "Import Table File",
                "",
                "Table Files (*.xls *.xlsx)"
            )

            if not file_name:
                QMessageBox.information(self.parent, "Information", "导入已取消")
                return

            try:
                df = pd.read_excel(file_name)
            except Exception as e:
                QMessageBox.critical(self.parent, "Error", f"无法导入文件: {str(e)}")
                return

        # 特殊处理的列
        special_columns = ['总成绩', '课序号', '学分', '学时', '绩点', '等级成绩']

        def convert_to_float(value):
            if pd.isna(value):
                return -1.0
            if isinstance(value, str) and value.strip() == '合格':
                return -1.0
            try:
                return float(value)
            except ValueError:
                return -1.0

        # 处理所有列
        data = {}
        for col in df.columns:
            if col in special_columns:
                # 特殊处理这些列
                data[col] = df[col].apply(convert_to_float).tolist()
            else:
                # 其他列保持原样，但空值转为空字符串
                data[col] = df[col].fillna('').astype(str).tolist()

        # 根据数据来源决定是否转置数据
        if not self.file_from_scraper:
            transposed_data = [dict(zip(data.keys(), row)) for row in zip(*data.values())]
        else:
            transposed_data = df.to_dict('records')

        # 对特殊列进行排序
        # for col in special_columns:
        #     if col in data:
        #         transposed_data.sort(key=lambda x: x.get(col, -1), reverse=True)

        # 创建用户信息字典
        user_info = {
            "姓名": name,
            "学号": student_id
        }

        # 将用户信息添加到数据的开头
        transposed_data.insert(0, user_info)

        # 确保 '../data' 目录存在
        os.makedirs(data_dir, exist_ok=True)

        with open(json_file_name, 'w', encoding='utf-8') as f:
            json.dump(transposed_data, f, ensure_ascii=False, indent=4)

        # 显示成功消息
        success_message = f"成功导入文件并保存为JSON。\n保存位置：{json_file_name}\n导入的列：{', '.join(df.columns)}\n特殊处理的列：{', '.join(special_columns)}\n用户信息已添加到文件开头。"
        QMessageBox.information(self.parent, "Success", success_message)
