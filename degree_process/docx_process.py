import json
import os
import sys
from collections import namedtuple
from io import BytesIO
import re
from typing import Union

import pandas as pd
from PyQt6.QtWidgets import QFileDialog, QMessageBox
from docx import Document

sys.path.append(os.getcwd())


class DocxProcess:
    def __init__(self, parent=None):
        self.parent = parent
        self.required_column = "理论教学学时"
        self.config_path = os.path.join("..", "config", "user_config.json")
        self.last_file_path = self.load_last_file_path()

    def load_last_file_path(self):
        if os.path.exists(self.config_path):
            with open(self.config_path, 'r') as f:
                config = json.load(f)
                return config.get('education_program_file_path', '')
        return ''

    def save_last_file_path(self, file_path):
        os.makedirs(os.path.dirname(self.config_path), exist_ok=True)

        # 读取现有的 JSON 内容
        existing_data = {}
        if os.path.exists(self.config_path):
            try:
                with open(self.config_path, 'r') as f:
                    existing_data = json.load(f)
            except json.JSONDecodeError:
                # 如果文件存在但不是有效的 JSON，我们就使用空字典
                pass

        # 更新 education_program_file_path
        existing_data['education_program_file_path'] = file_path

        # 写回文件，保持原有格式
        with open(self.config_path, 'w') as f:
            json.dump(existing_data, f, indent=4, ensure_ascii=False)

    def import_docx(self):
        if self.last_file_path and os.path.exists(self.last_file_path):
            reply = QMessageBox.question(self.parent, '使用上次文件',
                                         f"是否使用上次导入的文件?\n{self.last_file_path}",
                                         QMessageBox.StandardButton.Yes |
                                         QMessageBox.StandardButton.No,
                                         QMessageBox.StandardButton.No)
            if reply == QMessageBox.StandardButton.Yes:
                return self.process_file(self.last_file_path)

        file_name, _ = QFileDialog.getOpenFileName(
            self.parent,
            "导入 Word 文档",
            "",
            "Word 文档 (*.docx)"
        )

        if not file_name:
            QMessageBox.information(self.parent, "提示", "导入已取消")
            return None

        if not file_name.endswith('.docx'):
            QMessageBox.warning(self.parent, "警告", "请选择 .docx 格式的文件！")
            return None

        self.save_last_file_path(file_name)
        return self.process_file(file_name)

    def process_file(self, file_name):
        try:
            with open(file_name, 'rb') as file:
                docx_content = BytesIO(file.read())

            document = Document(docx_content)
            tables_with_paragraphs = self.extract_tables_and_paragraphs(document)
            self.export_to_json(results=tables_with_paragraphs)  # 将结果保存到json文件

            QMessageBox.information(self.parent, "成功", f"成功导入文件: {file_name}")

            return tables_with_paragraphs

        except Exception as e:
            QMessageBox.critical(self.parent, "错误", f"导入文件时发生错误: {str(e)}")
            return None

    def extract_credit_info(self, strings):
        """
        :param strings: 培养方案中每类课程的最低修读学分数
        :return: 格式化后的列表
        """
        # 定义一个命名元组来存储每类课程的信息
        CourseInfo = namedtuple('CourseInfo', ['course_type', 'required_credits', 'elective_credits'])

        # 初始化存储结果的列表
        credit_info = []

        # 定义正则表达式模式
        pattern = r'(.*?)\s*最低必修学分数[:：]\s*(\d+)\s*最低选修学分数[:：]\s*(\d+)'

        # 用于检查重复的集合
        seen_course_types = set()

        for string in strings:
            # 使用正则表达式匹配
            match = re.search(pattern, string)
            if match:
                course_type = match.group(1).strip()
                required_credits = int(match.group(2))
                elective_credits = int(match.group(3))

                # 如果这个课程类型还没有被记录，则添加信息
                if course_type not in seen_course_types:
                    credit_info.append(CourseInfo(course_type, required_credits, elective_credits))
                    seen_course_types.add(course_type)
                # 如果已经存在，则直接忽略

        return credit_info

    def extract_tables_and_paragraphs(self, document):
        results = []
        paragraphs = list(document.paragraphs)
        tables = list(document.tables)

        text = [p.text for p in paragraphs]
        relevant_paragraph = [content for content in text if
                              any(keyword in content for keyword in ["最低选修学分数", "最低必修学分数"])]
        # 获取每类课程需要的最低学分数
        score_need = self.extract_credit_info(relevant_paragraph)

        # 删除不需要的表格
        delete_index = []
        for i, table in enumerate(tables):
            if not table.rows:
                delete_index.append(i)
                continue

            header_row = [cell.text.strip() for cell in table.rows[0].cells]
            if self.required_column not in header_row:
                delete_index.append(i)

        # 使用列表推导式删除标记的项目
        tables = [table for i, table in enumerate(tables) if i not in delete_index]

        for i, table in enumerate(tables):
            # 提取表格数据
            header_row = [cell.text.strip() for cell in table.rows[0].cells]
            table_data = [
                [cell.text.strip() for cell in row.cells]
                for row in table.rows[1:]
            ]

            # 打印表格的基本信息
            print(f"表格大小: {len(table_data)} 行 x {len(header_row)} 列")
            print("列名:", ", ".join(header_row))

            results.append({
                'table': {
                    'header': header_row,
                    'data': table_data
                },
                'info': score_need[i] if i < len(score_need) else None
            })

        return results

    def export_to_json(self, results, json_filename="degree_progress.json"):
        """
        :param json_filename: 存储json的文件名
        :param results: self.extract_tables_and_paragraphs的返回值
        """
        # 确保 ../config/ 目录存在
        config_dir = os.path.abspath(os.path.join(os.path.dirname(__file__), '..', 'config'))
        os.makedirs(config_dir, exist_ok=True)

        # 构建完整的文件路径
        file_path = os.path.join(config_dir, json_filename)

        # 将结果转换为可序列化的格式
        serializable_results = []
        for item in results:
            serializable_item = {
                'table': item['table'],  # 这里的 'table' 已经是一个字典，包含 'header' 和 'data'
                'info': item['info']
            }
            serializable_results.append(serializable_item)

        # 将结果写入 JSON 文件
        with open(file_path, 'w', encoding='utf-8') as f:
            json.dump(serializable_results, f, ensure_ascii=False, indent=4)

        print(f"Results saved to {file_path}")

    def import_from_json(self, json_filename="degree_progress.json"):
        config_dir = os.path.abspath(os.path.join(os.path.dirname(__file__), '..', 'config'))
        file_path = os.path.join(config_dir, json_filename)

        # 检查文件是否存在
        if not os.path.exists(file_path):
            print(f"File not found: {file_path}")
            return None

        # 从 JSON 文件读取数据
        with open(file_path, 'r', encoding='utf-8') as f:
            loaded_results = json.load(f)

        # 转换回原始格式
        results = []
        for item in loaded_results:
            result = {
                'table': item['table'],  # 这里的 'table' 已经是一个字典，包含 'header' 和 'data'
                'info': item['info']
            }
            results.append(result)

        print(f"Results loaded from {file_path}")
        return results
