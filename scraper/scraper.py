import hashlib
from html import unescape
from xml.sax import parseString

from selenium import webdriver
from selenium.webdriver.chrome.options import Options as ChromeOptions
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException, NoSuchElementException
import time
import os
import pandas as pd


def open_browser_and_navigate(url):
    chrome_options = ChromeOptions()
    chrome_options.add_argument("--start-maximized")
    driver = webdriver.Chrome(options=chrome_options)

    driver.get(url)
    print(f"浏览器已打开并导航到指定URL。请手动完成登录操作，然后按回车键继续...")
    input()  # 等待用户输入

    return driver


def scrape_and_save_data(driver, valid_column_headers: list):
    try:
        WebDriverWait(driver, 20).until(
            EC.presence_of_element_located((By.XPATH, '//*[starts-with(@id, "contentqb-index-table-")]'))
        )

        content_elements = driver.find_elements(By.XPATH, '//*[starts-with(@id, "contentqb-index-table-")]')

        print(f"找到 {len(content_elements)} 个符合条件的元素")

        if not os.path.exists("table_contents"):
            os.makedirs("table_contents")

        for i, content_element in enumerate(content_elements, 1):
            try:
                column_header_element = content_element.find_element(By.XPATH,
                                                                     './/*[starts-with(@id, "columntableqb-index-table-")]')
                column_headers = column_header_element.find_elements(By.XPATH, './/div[@role="columnheader"]')

                column_titles = [header.find_element(By.TAG_NAME, 'span').get_attribute('title') for header in
                                 column_headers]

                # 只保留有效的列
                valid_indices = [i for i, title in enumerate(column_titles) if title in valid_column_headers]
                valid_titles = [title for title in column_titles if title in valid_column_headers]

                content_table_element = content_element.find_element(By.XPATH,
                                                                     './/*[starts-with(@id, "contenttableqb-index-table-")]')
                tbody_element = content_table_element.find_element(By.TAG_NAME, 'tbody')
                content_rows = tbody_element.find_elements(By.TAG_NAME, 'tr')

                data = []
                for row in content_rows:
                    cells = row.find_elements(By.TAG_NAME, 'td')
                    row_data = []
                    for i in valid_indices:
                        if i < len(cells):
                            a_element = cells[i].find_element(By.TAG_NAME, 'a')
                            row_data.append(a_element.text.strip())
                        else:
                            row_data.append('')
                    data.append(row_data)

                df = pd.DataFrame(data, columns=valid_titles)
                file_path = f"table_contents/table_content_{i}.xlsx"
                df.to_excel(file_path, index=False)

                print(f"表格 {i} 的内容已保存到 {file_path}")

            except NoSuchElementException as e:
                print(f"处理元素 {i} 时没有找到指定的元素: {e}")
            except Exception as e:
                print(f"处理元素 {i} 时发生错误: {e}")

    except TimeoutException:
        print("等待表格加载超时")
    except Exception as e:
        print(f"发生错误: {e}")


def get_column_titles(driver):
    try:
        content_elements = driver.find_elements(By.XPATH, '//*[@id="contentqb-index-table-"]')
        all_titles = []

        for content in content_elements:
            header_elements = content.find_elements(By.XPATH, './/div[@role="columnheader"]')
            titles = []
            for header in header_elements:
                span = header.find_element(By.TAG_NAME, 'span')
                title = span.get_attribute('title')
                titles.append(title)

            all_titles.append(titles)

        print(f"找到 {len(all_titles)} 个表格的列标题")
        for i, titles in enumerate(all_titles, 1):
            print(f"表格 {i}: 找到 {len(titles)} 个列标题")

        return all_titles
    except Exception as e:
        print(f"获取列标题时发生错误: {e}")
        return []


def main():
    url = "https://jw.xmu.edu.cn/jwapp/sys/cjcx/*default/index.do?t_s=1723166960886&amp_sec_version_=1&gid_=SXBVK1NhazRDMGZOSHpjMWVFSmhUNGJ1ZFRJUGxaRUxpbGpiTHRNZVYyQ044U0VjRi9BcmZCVzdlek5YL25oZHMzeFU2eEZpVWlEcDJ0L3F1Q3ZxL2c9PQ&EMAP_LANG=zh&THEME=cherry#/cjcx"
    driver = open_browser_and_navigate(url)
    column_headers = ["操作", "学年学期", "课程名", "课程号", "总成绩", "课序号", "课程类别", "课程性质", "学分",
                      "学时", "修读方式", "是否主修", "考试日期", "绩点", "重修重考", "等级成绩类型", "考试类型",
                      "开课单位", "是否及格", "是否有效", "特殊原因"]

    try:
        scrape_and_save_data(driver, column_headers)
    finally:
        time.sleep(5)
        driver.quit()


if __name__ == "__main__":
    main()
