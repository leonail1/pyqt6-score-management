import os
import time
import pandas as pd
from PyQt6.QtCore import QTimer, Qt
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

from PyQt6.QtWidgets import QApplication, QDialog, QPushButton, QVBoxLayout, QLabel, QHBoxLayout, QMessageBox, \
    QLineEdit, QWidget
import sys


class WelcomePage(QWidget):
    def __init__(self, default_username="", default_password=""):
        super().__init__()
        self.scraper = GradeScraper()
        self.initUI(default_username, default_password)

    def initUI(self, default_username="", default_password=""):
        self.setWindowTitle('成绩爬虫欢迎页面')
        self.setGeometry(300, 300, 300, 250)

        layout = QVBoxLayout()

        welcome_label = QLabel('欢迎使用成绩爬虫程序', self)
        welcome_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        layout.addWidget(welcome_label)

        self.username_input = QLineEdit(self)
        self.username_input.setPlaceholderText("请输入学号/工号")
        self.username_input.setText(default_username)
        layout.addWidget(self.username_input)

        self.password_input = QLineEdit(self)
        self.password_input.setPlaceholderText("请输入密码")
        self.password_input.setEchoMode(QLineEdit.EchoMode.Password)
        self.password_input.setText(default_password)
        layout.addWidget(self.password_input)

        button_layout = QHBoxLayout()

        start_button = QPushButton('开始爬虫', self)
        start_button.clicked.connect(self.start_scraping)
        button_layout.addWidget(start_button)

        cancel_button = QPushButton('取消', self)
        cancel_button.clicked.connect(self.close)
        button_layout.addWidget(cancel_button)

        layout.addLayout(button_layout)

        self.setLayout(layout)

    def start_scraping(self):
        username = self.username_input.text()
        password = self.password_input.text()

        if not username or not password:
            QMessageBox.warning(self, '错误', '请输入用户名和密码')
            return

        reply = QMessageBox.question(self, '确认', '确定要开始爬虫操作吗？',
                                     QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
                                     QMessageBox.StandardButton.No)

        if reply == QMessageBox.StandardButton.Yes:
            self.close()
            self.scraper.main(default_username=username, default_password=password)
        else:
            QMessageBox.information(self, '取消', '爬虫操作已取消')


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


class CredentialsDialog(QDialog):
    def __init__(self, default_username="", default_password=""):
        super().__init__()
        self.setWindowTitle("登录信息")
        self.setGeometry(100, 100, 300, 150)

        layout = QVBoxLayout()

        self.username_input = QLineEdit(self)
        self.username_input.setPlaceholderText("请输入学号/工号")
        self.username_input.setText(default_username)
        layout.addWidget(self.username_input)

        self.password_input = QLineEdit(self)
        self.password_input.setPlaceholderText("请输入密码")
        self.password_input.setEchoMode(QLineEdit.EchoMode.Password)
        self.password_input.setText(default_password)
        layout.addWidget(self.password_input)

        button_layout = QHBoxLayout()

        login_button = QPushButton("登录")
        login_button.clicked.connect(self.accept)
        button_layout.addWidget(login_button)

        cancel_button = QPushButton("取消")
        cancel_button.clicked.connect(self.reject)
        button_layout.addWidget(cancel_button)

        layout.addLayout(button_layout)

        self.setLayout(layout)

    def get_credentials(self):
        return self.username_input.text(), self.password_input.text()


class TimedMessageBox(QMessageBox):
    def __init__(self, timeout=3000, *args, **kwargs):
        super(TimedMessageBox, self).__init__(*args, **kwargs)
        self.timeout = timeout
        self.timer = QTimer(self)
        self.timer.timeout.connect(self.close)
        self.timer.start(self.timeout)

    def showEvent(self, event):
        self.timer.start(self.timeout)
        super(TimedMessageBox, self).showEvent(event)

    def closeEvent(self, event):
        self.timer.stop()
        super(TimedMessageBox, self).closeEvent(event)


class GradeScraper:
    def __init__(self):
        self.valid_column_headers = [
            "学年学期", "课程名", "课程号", "总成绩", "课序号", "课程类别", "课程性质", "学分",
            "学时", "修读方式", "是否主修", "考试日期", "绩点", "重修重考", "等级成绩类型", "考试类型",
            "开课单位", "是否及格", "是否有效"
        ]
        self.driver = None
        self.app = QApplication.instance() or QApplication(sys.argv)

    def show_message(self, title, message, timeout=3000):
        msg_box = TimedMessageBox(timeout=timeout, icon=QMessageBox.Icon.Information, text=message, windowTitle=title)
        msg_box.exec()

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

    def input_credentials(self, default_username="", default_password=""):
        # dialog = CredentialsDialog(default_username, default_password)
        # if dialog.exec() == QDialog.DialogCode.Accepted:
        #     username, password = dialog.get_credentials()
        # else:
        #     self.show_message("操作取消", "登录操作被取消")
        #     return False

        username, password = default_username, default_password
        try:
            # 定位并输入用户名
            username_input = WebDriverWait(self.driver, 10).until(
                EC.presence_of_element_located((By.ID, "username"))
            )
            username_input.clear()
            username_input.send_keys(username)

            # 定位并输入密码
            password_input = WebDriverWait(self.driver, 10).until(
                EC.presence_of_element_located((By.ID, "password"))
            )
            password_input.clear()
            password_input.send_keys(password)

            # 定位并点击登录按钮
            login_button = WebDriverWait(self.driver, 10).until(
                EC.element_to_be_clickable((By.ID, "login_submit"))
            )
            login_button.click()

            # self.show_message("登录操作完成", "成功输入用户名和密码并点击登录按钮")
            return True
        except Exception as e:
            self.show_message("登录失败", f"无法完成登录操作: {str(e)}")
            return False

    def wait_for_element(self, by, value, timeout=20, element=None):
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

                        valid_indices = [i for i, title in enumerate(column_titles) if
                                         title in self.valid_column_headers]
                        valid_titles = [title for title in column_titles if title in self.valid_column_headers]

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

    def click_login_button(self):
        try:
            locators = [
                (By.ID, "userNameLogin_a"),
                (By.LINK_TEXT, "账号登录"),
                (By.CLASS_NAME, "loginFont_a"),
                (By.XPATH, "//a[contains(text(),'账号登录')]"),
                (By.CSS_SELECTOR, "a.loginFont_a#userNameLogin_a")
            ]

            for locator in locators:
                try:
                    button = WebDriverWait(self.driver, 5).until(
                        EC.element_to_be_clickable(locator)
                    )
                    button.click()
                    # self.show_message("操作成功", "成功点击了\"账号登录\"按钮")
                    return
                except:
                    continue

            raise Exception("无法找到或点击\"账号登录\"按钮")

        except Exception as e:
            self.show_message("操作失败", f"无法点击\"账号登录\"按钮: {str(e)}")

    def run(self, url, browser_type='chrome', default_username="", default_password=""):
        self.open_browser_and_navigate(url, browser_type)
        try:
            self.click_login_button()  # 点击"账号登录"按钮
            if self.input_credentials(default_username, default_password):
                self.show_message("提示", "请在20秒之内点击页面上的\"全部成绩\"按钮")
                self.scrape_and_save_data()
            else:
                self.show_message("程序结束", "操作已取消，程序结束。")
        finally:
            if self.driver:
                self.driver.quit()

    def main(self, default_username="", default_password=""):
        url = ("https://jw.xmu.edu.cn/jwapp/sys/cjcx/*default/index.do?t_s=1723166960886&amp_sec_version_=1&gid_"
               "=SXBVK1NhazRDMGZOSHpjMWVFSmhUNGJ1ZFRJUGxaRUxpbGpiTHRNZVYyQ044U0VjRi9BcmZCVzdlek5YL25oZHMzeFU2eEZpVWlEcDJ0L3F1Q3ZxL2c9PQ&EMAP_LANG=zh&THEME=cherry#/cjcx")
        scraper = GradeScraper()
        scraper.run(url, default_username=default_username, default_password=default_password)


def start():
    app = QApplication(sys.argv)

    welcome = WelcomePage(default_username="37220222203691", default_password="mudwa2-kihjar-wipjiF")
    welcome.show()

    sys.exit(app.exec())


if __name__ == "__main__":
    start()
