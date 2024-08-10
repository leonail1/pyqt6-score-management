import json
import sys
import os
from typing import Dict, Union, List

sys.path.append(os.getcwd())


class StudentScoreAnalyzer():
    def __init__(self, parent=None):
        self.parent = parent
        self.score_data = None

    def load_score_data(self, student_id: str) -> Dict[str, Union[Dict[str, str], List[Dict[str, Union[str, float]]]]]:
        data_dir = os.path.abspath(os.path.join(os.path.dirname(__file__), '..', 'data'))
        file_path = os.path.join(data_dir, f"{student_id}.json")

        try:
            with open(file_path, 'r', encoding='utf-8') as file:
                data = json.load(file)

            # 验证数据结构
            if not isinstance(data, list) or len(data) < 2:
                raise ValueError("Invalid data structure")

            # 提取学生信息和成绩信息
            student_info = data[0]
            scores = data[1:]

            # 验证学生信息
            if "姓名" not in student_info or "学号" not in student_info:
                raise ValueError("Missing student information")

            self.score_data = {
                "student_info": student_info,
                "scores": scores
            }

            return self.score_data

        except FileNotFoundError:
            print(f"File not found: {file_path}")
            return None
        except json.JSONDecodeError:
            print(f"Invalid JSON in file: {file_path}")
            return None
        except ValueError as e:
            print(f"Data validation error: {str(e)}")
            return None
        except Exception as e:
            print(f"An error occurred while loading data: {str(e)}")
            return None

    import os
    import json
    from typing import Dict, Union, List

    def save_score_data(self, score_data, student_id) -> bool:
        data_dir = os.path.abspath(os.path.join(os.path.dirname(__file__), '..', 'data'))
        file_path = os.path.join(data_dir, f"{student_id}.json")

        try:
            # 构造要保存的数据结构
            data_to_save = [self.score_data["student_info"]] + self.score_data["scores"]

            # 确保目录存在
            os.makedirs(os.path.dirname(file_path), exist_ok=True)

            # 写入文件
            with open(file_path, 'w', encoding='utf-8') as file:
                json.dump(data_to_save, file, ensure_ascii=False, indent=4)

            print(f"Data successfully saved to {file_path}")
            return True

        except Exception as e:
            print(f"An error occurred while saving data: {str(e)}")
            return False
