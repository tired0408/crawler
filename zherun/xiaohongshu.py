"""
抓取小红书超自然行动组的信息
"""
import os
import sys
root_path = os.path.dirname(os.path.dirname(os.path.dirname(__file__)))
sys.path.append(root_path)
import datetime
import pandas as pd
from crawler.utils import init_chrome
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC


chromedriver_path = os.path.join(root_path, r"chromedriver_mac_arm64_114\chromedriver.exe")
chrome_path = os.path.join(root_path, r"chromedriver_mac_arm64_114\chrome114\App\Chrome-bin\chrome.exe")
user_path = r'C:\Users\Administrator\AppData\Local\Google\Chrome\crawler'
url = r"https://www.xiaohongshu.com/explore"
data_path = os.path.join(root_path, "crawler/zherun/data.xlsx")
print("初始化爬虫抓取工具")
driver = init_chrome(chromedriver_path, chrome_path=chrome_path, user_path=user_path, headless=True)
print("打开网页,开始抓取数据")
driver.get(url)
ele = WebDriverWait(driver, 60).until(EC.element_to_be_clickable((By.XPATH, "//input[@id='search-input']")))
ele.send_keys("超自然行动组")
ele = WebDriverWait(driver, 60).until(EC.element_to_be_clickable((By.XPATH, "//div[@class='search-icon']")))
ele.click()
ele = WebDriverWait(driver, 60).until(EC.element_to_be_clickable((By.XPATH, "//div[@class='onebox']/a")))
ele.click()
print("进入新标签页")
WebDriverWait(driver, 60).until(EC.number_of_windows_to_be(2))
driver.switch_to.window(driver.window_handles[-1])
ele = WebDriverWait(driver, 60).until(EC.element_to_be_clickable((By.XPATH, "//div[@class='user-interactions']/div/span[text()='粉丝']/preceding-sibling::span[1]")))
fans = float(ele.text[:-1])
ele = WebDriverWait(driver, 60).until(EC.element_to_be_clickable((By.XPATH, "//div[@class='user-interactions']/div/span[text()='获赞与收藏']/preceding-sibling::span[1]")))
like = float(ele.text[:-1])
print("关闭浏览器")
driver.quit()
now_time = datetime.datetime.now()
print(f"[{now_time}]今日粉丝量:{fans}, 获赞与收藏:{like}")
print("获取历史数据,并新增当前数据内容")
df = pd.read_excel(data_path, header=0)
fans_growth = fans - df.iloc[-1, 1]
like_growth = like - df.iloc[-1, 2]
df.loc[len(df)] = [now_time, fans, like, fans_growth, like_growth]
with pd.ExcelWriter(data_path, engine="openpyxl") as writer:
    df.to_excel(writer, index=False, sheet_name="Sheet1")
    worksheet = writer.sheets["Sheet1"]
    worksheet.column_dimensions['A'].width = 21  # 第一列
    worksheet.column_dimensions['B'].width = 18
    worksheet.column_dimensions['C'].width = 18
    worksheet.column_dimensions['D'].width = 18
    worksheet.column_dimensions['E'].width = 18
print("数据拉取完成")

