"""
爱医助医的网站数据分析脚本1
"""
import os
import time
import shutil
import pandas as pd
from utils import init_chrome
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

print(__file__)
# acount = "18611756193"
# password = "secret"
# date_range = "2025.02-2025.04"
# region = "xxxx"

# search_info = [["2025", "02", "其他区域"], ["2025", "01", "其他区域"]]

# path = r"E:\NewFolder\love_doctor"
# download_path = r"E:\NewFolder\love_doctor\file"
# if os.path.exists(download_path):
#     shutil.rmtree(download_path)
# os.makedirs(download_path)
# chromedriver_path = os.path.join(path, r"..\chromedriver_mac_arm64_114\chromedriver.exe")
# chrome_path = os.path.join(path, r"..\chromedriver_mac_arm64_114\chrome114\App\Chrome-bin\chrome.exe")
# driver = init_chrome(chromedriver_path, download_path, chrome_path=chrome_path, is_proxy=False)
# pattern = (By.XPATH, "//button[contains(text(), '登录')]")
# print("打开网页")
# driver.get(r"http://cbs.aiyizhuyi.com/ayzy/cbs/yearly_data/monthsum")
# WebDriverWait(driver, 10).until(EC.element_to_be_clickable(pattern))
# print("输入账号密码")
# ele = driver.find_element(By.NAME, "login")
# ele.clear()
# ele.send_keys(acount)
# ele = driver.find_element(By.NAME, "password")
# ele.clear()
# ele.send_keys(password)
# print("登录系统")
# ele = driver.find_element(*pattern)
# ele.click()
# print("点击《数据报告》模块")
# pattern = (By.XPATH, "//div[contains(@id, 'layout-canvas')]//ul[contains(@class, 'mainmenu-nav')]//span[contains(text(), '数据报告')]")
# ele = WebDriverWait(driver, 10).until(EC.element_to_be_clickable(pattern))
# ele.click()
# pattern = (By.XPATH, "//span[contains(text(), '设备耗材商务系统展示')]")
# ele = WebDriverWait(driver, 10).until(EC.presence_of_element_located(pattern))
# print("滚动到《设备耗材商务系统展示》选择标签")
# driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", ele)
# ele = WebDriverWait(driver, 10).until(EC.element_to_be_clickable(pattern))
# print("点击《设备耗材商务系统展示》")
# ele.click()
# for year, month, region in search_info:
#     print(f"查看当前是否为目标年份:{year}")
#     ele = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.ID, "select2-year-container")))
#     if year not in ele.text:
#         print("年份有误,选择年份")
#         ele.click()
#         ele = WebDriverWait(driver, 10).until(EC.visibility_of_element_located((By.XPATH, f"//ul[@id='select2-year-results']//li[contains(text(), f'{year}')]")))
#         ele.click()
#         WebDriverWait(driver, 10).until(EC.text_to_be_present_in_element((By.ID, "select2-month-container"), f"{year}年"))
#     print(f"查看当前是否为目标月份:{month}")
#     ele = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.ID, "select2-month-container")))
#     if month not in ele.text:
#         print("月份有误,选择月份")
#         ele.click()
#         pattern = (By.XPATH, f"//ul[@id='select2-month-results']//li[contains(text(), '{int(month)}')]")
#         ele = WebDriverWait(driver, 10).until(EC.presence_of_element_located(pattern))
#         driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", ele)
#         ele = WebDriverWait(driver, 10).until(EC.element_to_be_clickable(pattern))
#         ele.click()
#         WebDriverWait(driver, 10).until(EC.text_to_be_present_in_element((By.ID, "select2-month-container"), f"{int(month)}月"))
#     print(f"查看当前区域是否为目标区域:{region}")
#     ele = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.ID, "select2-region-container")))
#     if region not in ele.text:
#         print("区域有误,选择区域")
#         ele.click()
#         pattern = (By.XPATH, f"//ul[@id='select2-region-results']//li[contains(text(), f'{region}')]")
#         ele = WebDriverWait(driver, 10).until(EC.presence_of_element_located(pattern))
#         driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", ele)
#         ele = WebDriverWait(driver, 10).until(EC.element_to_be_clickable(pattern))
#         ele.click()
#         WebDriverWait(driver, 10).until(EC.text_to_be_present_in_element((By.ID, "select2-region-container"), f"{region}"))
#     print("导出EXCEL到目标文件夹")
#     ele = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.ID, "download1")))
#     ele.click()
#     st = time.time()
#     print("等待文件下载完成")
#     name = f"{year}-{month}_{region}_全部设备耗材汇总.xlsx"
#     file_path = os.path.join(download_path, name)
#     while time.time() - st < 300:
#         if os.path.exists(file_path):
#             break
#         time.sleep(5)
#     print(f"{name}, 已下载完成")
# print("关闭浏览器")
# driver.quit()
# print("整理下载后的EXCEL文件,并合并到《设备耗材汇总》表")
# output_path = os.path.join(path, "设备耗材汇总.xlsx")
# if not os.path.exists(output_path):
#     raise Exception("《设备耗材汇总》模板文件不存在")
# print("读取目标表数据")
# output_df = pd.read_excel(output_path)
# print("将导出的数据合并到目标表中")
# for file in os.listdir(download_path):
#     file_path = os.path.join(download_path, file)
#     input_df = pd.read_excel(file_path)
#     input_df = input_df.iloc[:-1]
#     output_df = pd.concat([output_df, input_df], ignore_index=True)
# print("删除重复数据")
# output_df = output_df.drop_duplicates()
# print("保存数据")
# output_df.to_excel(output_path, index=False)
# print("《设备耗材汇总》表已完成")