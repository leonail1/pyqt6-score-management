import os
import time
import pandas as pd
from selenium import webdriver
from selenium.common.exceptions import TimeoutException
from selenium.webdriver.chrome.options import Options as ChromeOptions
from selenium.webdriver.chrome.service import Service as ChromeService
from selenium.webdriver.common.by import By
from selenium.webdriver.edge.options import Options as EdgeOptions
from selenium.webdriver.edge.service import Service as EdgeService
from selenium.webdriver.safari.options import Options as SafariOptions
from selenium.webdriver.safari.service import Service as SafariService
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait
from webdriver_manager.chrome import ChromeDriverManager
from webdriver_manager.microsoft import EdgeChromiumDriverManager

from PyQt6.QtWidgets import QApplication, QDialog, QPushButton, QVBoxLayout, QLabel, QHBoxLayout, QMessageBox
import sys


class LoginDialog(QDialog):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("操作确认")
        self.setGeometry(100, 100, 400, 150)

        layout = QVBoxLayout()

        label = QLabel("请导航到成绩-全部成绩页面，然后点击确认按钮继续。")
        label.setWordWrap(True)
        layout.addWidget(label)

        button_layout = QHBoxLayout()

        confirm_button = QPushButton("确认")
        confirm_button.clicked.connect(self.accept)
        button_layout.addWidget(confirm_button)

        cancel_button = QPushButton("取消")
        cancel_button.clicked.connect(self.reject)
        button_layout.addWidget(cancel_button)

        layout.addLayout(button_layout)

        self.setLayout(layout)


class GradeScraper:
    def __init__(self):
        self.driver = None
        self.column_headers = [
            "学年学期", "课程名", "课程号", "总成绩", "课序号", "课程类别", "课程性质", "学分",
            "学时", "修读方式", "是否主修", "考试日期", "绩点", "重修重考", "等级成绩类型", "考试类型",
            "开课单位", "是否及格", "是否有效", "特殊原因"
        ]
        self.app = QApplication.instance() or QApplication(sys.argv)

    def show_message(self, title, message):
        QMessageBox.information(None, title, message)

    def open_browser_and_navigate(self, url, browser_type='chrome'):
        browser_type = browser_type.lower()

        if browser_type == 'chrome':
            options = ChromeOptions()
            options.add_argument("--start-maximized")
            service = ChromeService(ChromeDriverManager().install())
            self.driver = webdriver.Chrome(service=service, options=options)
        elif browser_type == 'edge':
            options = EdgeOptions()
            options.add_argument("--start-maximized")
            service = EdgeService(EdgeChromiumDriverManager().install())
            self.driver = webdriver.Edge(service=service, options=options)
        elif browser_type == 'safari':
            options = SafariOptions()
            service = SafariService()
            self.driver = webdriver.Safari(service=service, options=options)
        else:
            raise ValueError("不支持的浏览器类型。请选择 'chrome'、'edge' 或 'safari'。")

        self.driver.get(url)
        # self.show_message("浏览器已启动", f"{browser_type.capitalize()} 浏览器已打开并导航到指定URL。")

        dialog = LoginDialog()
        result = dialog.exec()

        if result == QDialog.DialogCode.Rejected:
            self.show_message("操作取消", "操作被用户取消")
            self.driver.quit()
            return False

        return True

    def wait_for_element(self, by, value, timeout=10, element=None):
        if element:
            return WebDriverWait(element, timeout).until(
                EC.presence_of_element_located((by, value))
            )
        return WebDriverWait(self.driver, timeout).until(
            EC.presence_of_element_located((by, value))
        )

    @staticmethod
    def get_element_text(element):
        return element.get_attribute('textContent').strip()

    def scrape_and_save_data(self):
        try:
            self.wait_for_element(By.XPATH, '//*[starts-with(@id, "contentqb-index-table-")]')
            content_elements = self.driver.find_elements(By.XPATH, '//*[starts-with(@id, "contentqb-index-table-")]')
            print(f"找到 {len(content_elements)} 个符合条件的元素")

            all_data = []
            directory = "table_contents"
            file_name = "all_tables_content.xlsx"
            file_path = os.path.join(os.path.dirname(os.path.realpath(__file__)), directory, file_name)

            # 确保目录存在
            os.makedirs(directory, exist_ok=True)

            # 如果文件已存在，删除它（覆盖写入）
            if os.path.exists(file_path):
                os.remove(file_path)
                print(f"已删除现有文件: {file_path}")

            with pd.ExcelWriter(file_path, engine='openpyxl') as writer:
                for i, content_element in enumerate(content_elements, 1):
                    try:
                        column_header_element = self.wait_for_element(By.XPATH,
                                                                      './/*[starts-with(@id, "columntableqb-index-table-")]',
                                                                      element=content_element)
                        column_headers = column_header_element.find_elements(By.XPATH, './/div[@role="columnheader"]')
                        column_titles = [self.get_element_text(header.find_element(By.TAG_NAME, 'span')) for header in
                                         column_headers]

                        valid_indices = [i for i, title in enumerate(column_titles) if title in self.column_headers]
                        valid_titles = [title for title in column_titles if title in self.column_headers]

                        content_table_element = self.wait_for_element(By.XPATH,
                                                                      './/*[starts-with(@id, "contenttableqb-index-table-")]',
                                                                      element=content_element)
                        tbody_element = self.wait_for_element(By.TAG_NAME, 'tbody', element=content_table_element)
                        content_rows = tbody_element.find_elements(By.TAG_NAME, 'tr')

                        data = []
                        for row in content_rows:
                            cells = row.find_elements(By.TAG_NAME, 'td')
                            row_data = []
                            for j in valid_indices:
                                if j < len(cells):
                                    cell_text = self.get_element_text(cells[j])
                                    row_data.append(cell_text)
                                else:
                                    row_data.append('')
                            data.append(row_data)

                        df = pd.DataFrame(data, columns=valid_titles)

                        # 删除重复列
                        df = df.T.drop_duplicates().T

                        sheet_name = df["学年学期"].iloc[
                            0] if "学年学期" in df.columns and not df.empty else f"Sheet_{i}"
                        df.to_excel(writer, sheet_name=sheet_name, index=False)

                        all_data.append(df)
                        print(f"表格 {i} 的内容已保存到工作表 '{sheet_name}'")

                    except Exception as e:
                        print(f"处理元素 {i} 时发生错误: {e}")

                if all_data:
                    # 创建总表
                    total_df = pd.concat(all_data, ignore_index=True)

                    # 删除总表中的重复列
                    total_df = total_df.T.drop_duplicates().T

                    total_df.to_excel(writer, sheet_name='总表', index=False)
                    print("所有数据已合并到'总表'工作表")
                else:
                    print("没有成功抓取到任何数据")

            self.show_message("保存成功", f"所有表格内容已保存到 {file_path}")

        except TimeoutException:
            self.show_message("错误", "等待表格加载超时")
        except Exception as e:
            self.show_message("错误", f"发生错误: {e}")

    def run(self, url, browser_type='chrome'):
        if self.open_browser_and_navigate(url, browser_type):
            try:
                self.scrape_and_save_data()
            finally:
                time.sleep(0.5)
                if self.driver:
                    self.driver.quit()
        else:
            self.show_message("程序结束", "操作已取消，程序结束。")


def main():
    url = ("https://jw.xmu.edu.cn/jwapp/sys/cjcx/*default/index.do?t_s=1723166960886&amp_sec_version_=1&gid_"
           "=SXBVK1NhazRDMGZOSHpjMWVFSmhUNGJ1ZFRJUGxaRUxpbGpiTHRNZVYyQ044U0VjRi9BcmZCVzdlek5YL25oZHMzeFU2eEZpVWlEcDJ0L3F1Q3ZxL2c9PQ&EMAP_LANG=zh&THEME=cherry#/cjcx")
    scraper = GradeScraper()
    scraper.run(url)


if __name__ == "__main__":
    main()
