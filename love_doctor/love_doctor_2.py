"""
爱医助医的网站数据分析脚本2
"""
import os
import time
import shutil
import datetime
import pandas as pd
from typing import Dict
from utils import init_chrome
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC



user_folder = r"E:\NewFolder\love_doctor"
output_name = "全部设备耗材流向跟踪表.xlsx"
print("读取设备耗材汇总表")
input_df = pd.read_excel(os.path.join(user_folder, "设备耗材汇总.xlsx"), header=[0, 1])
input_df.columns = [c1 if c2.startswith("Unnamed") else c2 for c1, c2 in input_df.columns]
input_df = input_df.sort_values(by="时间", ascending=True)
print("分析输入数据表")
output_dict: Dict[str, Dict] = {}
for _, row in input_df.iterrows():
    sum_data = row["本区域流向"] + row["跨区域流向"]
    key = f'{row["归属区域"]}_{row["省份"]}_{row["客户"]}'
    if key not in output_dict:
        output_dict[key] = {row["时间"]: sum_data}
    else:
        output_dict[key][row["时间"]] = sum_data
print("转换成《全部设备耗材流向跟踪表》所需的数据")
columns=["归属区域", "省份", "客户"]
now_time = datetime.datetime.now()
for i in range(1, 13):
    columns.append(f"{now_time.year-1}-{i:02d}")
columns.append(f"{now_time.year-1}年合计")
end_date = input_df.iloc[-1]["时间"]
end_time = datetime.datetime.strptime(end_date, "%Y-%m")
for i in range(1, end_time.month+1):
    columns.append(f"{end_time.year}-{i:02d}")
columns.append(f"{end_time.year}年合计")
columns.append("终端状态")
output_list = []
for key, value in output_dict.items():
    each_row = []
    region, province, custom = key.split("_")
    each_row.append(region)
    each_row.append(province)
    each_row.append(custom)
    for i in range(3, 15):
        each_row.append(value.get(columns[i], 0))
    each_row.append(sum(each_row[3:15]))
    end_index = len(columns) - 1
    for i in range(17, end_index):
        each_row.append(value.get(columns[i], 0))
    each_row.append(sum(each_row[17:end_index]))
    # 计算终端状态
    if each_row[16] <= 0 and each_row[-1] > 0:
        each_row.append("新增")
    elif each_row[16] > 0 and each_row[-1] > 0:
        each_row.append("存量")
    elif each_row[16] > 0 and each_row[-1] <= 0:
        each_row.append("丢失")
    else:
        each_row.append("异常")
print("构建《全部设备耗材流向跟踪表》")
output_df = pd.DataFrame(output_list, columns=columns)
output_df.to_excel(os.path.join(user_folder, output_name), index=False)
print("《全部设备耗材流向跟踪表》已输出")