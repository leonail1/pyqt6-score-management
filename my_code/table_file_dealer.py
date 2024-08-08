import json
import pandas as pd
import os
from PyQt6.QtWidgets import QFileDialog, QMessageBox, QInputDialog, QLineEdit

from .student_score_analyzer import StudentScoreAnalyzer
from my_window.StudentInfoWindow import StudentInfoWindow


class FileDealer:
    def __init__(self, parent):
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
                confirm = QMessageBox.question(self.parent, "确认覆盖",
                                               "确定要覆盖现有数据吗？这将删除原有数据。",
                                               QMessageBox.StandardButton.Yes |
                                               QMessageBox.StandardButton.No)
                if confirm == QMessageBox.StandardButton.Yes:
                    self.import_file(student_id_input=student_id)
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
                self.import_file(student_id_input=student_id)
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
        config_file = "pyqt6_score_management/config/user_config.json"
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
        # print(name, student_id)
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

        file_name, _ = QFileDialog.getOpenFileName(
            self.parent,
            "Import Table File",
            "",
            "Table Files (*.xls *.xlsx)"
        )

        if file_name:
            try:
                # 读取Excel文件
                df = pd.read_excel(file_name)

                # 定义必需的列名和可选的列名
                required_columns = ['课程名', '课程性质', '学分']
                optional_columns = ['学年学期', '等级成绩', '绩点']

                # 检查必需的列是否都存在
                missing_columns = [col for col in required_columns if col not in df.columns]
                if missing_columns:
                    raise ValueError(f"缺少必需的列：{', '.join(missing_columns)}")

                # 检查哪些可选的列在文件中存在
                found_optional_columns = [col for col in optional_columns if col in df.columns]

                # 所有要处理的列
                columns_to_process = required_columns + found_optional_columns

                # 从文件中提取数据并转置
                data = {}
                for col in columns_to_process:
                    if col == '等级成绩':
                        # 将等级成绩转换为字符串，空值转为空字符串
                        data[col] = df[col].fillna('').astype(str).tolist()
                    else:
                        # 其他列保持原样，但空值转为空字符串
                        data[col] = df[col].fillna('').tolist()

                transposed_data = [dict(zip(data.keys(), row)) for row in zip(*data.values())]

                # 创建用户信息字典
                user_info = {
                    "姓名": name,
                    "学号": student_id
                }

                # 将用户信息添加到转置数据的开头
                transposed_data.insert(0, user_info)

                # 确保 '../data' 目录存在
                os.makedirs(data_dir, exist_ok=True)

                with open(json_file_name, 'w', encoding='utf-8') as f:
                    json.dump(transposed_data, f, ensure_ascii=False, indent=4)

                # 显示成功消息
                success_message = f"成功导入文件并保存为JSON。\n保存位置：{json_file_name}\n必需列：{', '.join(required_columns)}\n可选列：{', '.join(found_optional_columns)}\n用户信息已添加到文件开头。"
                QMessageBox.information(self.parent, "Success", success_message)

            except Exception as e:
                # 显示错误消息
                QMessageBox.critical(self.parent, "Error", f"无法导入文件: {str(e)}")
        else:
            QMessageBox.information(self.parent, "Information", "导入已取消")
