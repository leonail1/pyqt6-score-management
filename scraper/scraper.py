import os
import time
import pandas as pd
from selenium import webdriver
from selenium.common.exceptions import TimeoutException, NoSuchElementException
from selenium.webdriver.chrome.options import Options as ChromeOptions
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait


def open_browser_and_navigate(url):
    chrome_options = ChromeOptions()
    chrome_options.add_argument("--start-maximized")
    driver = webdriver.Chrome(options=chrome_options)

    driver.get(url)
    print(f"浏览器已打开并导航到指定URL。请手动完成登录操作，然后按回车键继续...")
    input()  # 等待用户输入

    return driver


def wait_for_element(driver, by, value, timeout=20):
    return WebDriverWait(driver, timeout).until(
        EC.presence_of_element_located((by, value))
    )


def get_element_text(element):
    return element.get_attribute('textContent').strip()


def scrape_and_save_data(driver, valid_column_headers: list):
    try:
        wait_for_element(driver, By.XPATH, '//*[starts-with(@id, "contentqb-index-table-")]')
        content_elements = driver.find_elements(By.XPATH, '//*[starts-with(@id, "contentqb-index-table-")]')
        print(f"找到 {len(content_elements)} 个符合条件的元素")

        all_data = []
        directory = "table_contents"
        file_name = "all_tables_content.xlsx"
        file_path = os.path.join(os.path.realpath(__file__), directory, file_name)

        # 确保目录存在
        os.makedirs(directory, exist_ok=True)

        # 如果文件已存在，删除它（覆盖写入）
        if os.path.exists(file_path):
            os.remove(file_path)
            print(f"已删除现有文件: {file_path}")

        with pd.ExcelWriter(file_path, engine='openpyxl') as writer:
            for i, content_element in enumerate(content_elements, 1):
                try:
                    column_header_element = wait_for_element(content_element, By.XPATH,
                                                             './/*[starts-with(@id, "columntableqb-index-table-")]')
                    column_headers = column_header_element.find_elements(By.XPATH, './/div[@role="columnheader"]')
                    column_titles = [get_element_text(header.find_element(By.TAG_NAME, 'span')) for header in
                                     column_headers]

                    valid_indices = [i for i, title in enumerate(column_titles) if title in valid_column_headers]
                    valid_titles = [title for title in column_titles if title in valid_column_headers]

                    content_table_element = wait_for_element(content_element, By.XPATH,
                                                             './/*[starts-with(@id, "contenttableqb-index-table-")]')
                    tbody_element = wait_for_element(content_table_element, By.TAG_NAME, 'tbody')
                    content_rows = tbody_element.find_elements(By.TAG_NAME, 'tr')

                    data = []
                    for row in content_rows:
                        cells = row.find_elements(By.TAG_NAME, 'td')
                        row_data = []
                        for j in valid_indices:
                            if j < len(cells):
                                cell_text = get_element_text(cells[j])
                                row_data.append(cell_text)
                            else:
                                row_data.append('')
                        data.append(row_data)

                    df = pd.DataFrame(data, columns=valid_titles)

                    # 删除重复列
                    df = df.T.drop_duplicates().T

                    sheet_name = df["学年学期"].iloc[0] if "学年学期" in df.columns and not df.empty else f"Sheet_{i}"
                    df.to_excel(writer, sheet_name=sheet_name, index=False)

                    all_data.append(df)
                    print(f"表格 {i} 的内容已保存到工作表 '{sheet_name}'")

                except Exception as e:
                    print(f"处理元素 {i} 时发生错误: {e}")

            # 创建总表
            total_df = pd.concat(all_data, ignore_index=True)

            # 删除总表中的重复列
            total_df = total_df.T.drop_duplicates().T

            total_df.to_excel(writer, sheet_name='总表', index=False)
            print("所有数据已合并到'总表'工作表")

        print(f"所有表格内容已保存到 {file_path}")

    except TimeoutException:
        print("等待表格加载超时")
    except Exception as e:
        print(f"发生错误: {e}")


def main():
    url = ("https://jw.xmu.edu.cn/jwapp/sys/cjcx/*default/index.do?t_s=1723166960886&amp_sec_version_=1&gid_"
           "=SXBVK1NhazRDMGZOSHpjMWVFSmhUNGJ1ZFRJUGxaRUxpbGpiTHRNZVYyQ044U0VjRi9BcmZCVzdlek5YL25oZHMzeFU2eEZpVWlEcDJ0L3F1Q3ZxL2c9PQ&EMAP_LANG=zh&THEME=cherry#/cjcx")
    driver = open_browser_and_navigate(url)
    column_headers = ["学年学期", "课程名", "课程号", "总成绩", "课序号", "课程类别", "课程性质", "学分",
                      "学时", "修读方式", "是否主修", "考试日期", "绩点", "重修重考", "等级成绩类型", "考试类型",
                      "开课单位", "是否及格", "是否有效", "特殊原因"]

    try:
        scrape_and_save_data(driver, column_headers)
    finally:
        time.sleep(0.5)
        driver.quit()


if __name__ == "__main__":
    main()
