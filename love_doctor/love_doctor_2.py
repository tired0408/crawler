"""
爱医助医的网站数据分析脚本2
"""
import os
import time

import shutil
import datetime
import pandas as pd
from tqdm import tqdm
from openpyxl import load_workbook
from typing import Dict
from utils import read_multi_column, read_by_openpyxl
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC



user_folder = os.path.dirname(__file__)
output_name = "全部设备耗材流向跟踪表.xlsx"
print("读取设备耗材汇总表")
wb, ws, header = read_by_openpyxl(os.path.join(user_folder, "设备耗材汇总.xlsx"))
print("分析输入数据表")
output_dict: Dict[str, Dict] = {}
for row in tqdm(ws.iter_rows(min_row=3, values_only=True), total=ws.max_row-2, desc="分析中"):
    row_dict = dict(zip(header, row))
    sum_data = row_dict["本区域流向"] + row_dict["跨区域流向"]
    key = f'{row_dict["归属区域"]}_{row_dict["省份"]}_{row_dict["客户"]}'
    if key not in output_dict:
        output_dict[key] = {row_dict["时间"]: sum_data}
    else:
        output_dict[key][row_dict["时间"]] = sum_data
end_time = next(ws.iter_rows(min_row=ws.max_row, max_row=ws.max_row, values_only=True))
end_time = end_time[header.index("时间")]
end_time = datetime.datetime.strptime(end_time, "%Y-%m")
wb.close()
print("转换成《全部设备耗材流向跟踪表》所需的数据")
columns=["归属区域", "省份", "客户"]
now_time = datetime.datetime.now()
for i in range(1, 13):
    columns.append(f"{now_time.year-1}-{i:02d}")
columns.append(f"{now_time.year-1}年合计")
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
    end_index = len(columns) - 2
    for i in range(16, end_index):
        each_row.append(value.get(columns[i], 0))
    each_row.append(sum(each_row[16:end_index]))
    # 计算终端状态
    if each_row[16] <= 0 and each_row[-1] > 0:
        each_row.append("新增")
    elif each_row[16] > 0 and each_row[-1] > 0:
        each_row.append("存量")
    elif each_row[16] > 0 and each_row[-1] <= 0:
        each_row.append("丢失")
    else:
        each_row.append("未开发")
    output_list.append(each_row)
print("构建《全部设备耗材流向跟踪表》")
output_df = pd.DataFrame(output_list, columns=columns)
output_df.to_excel(os.path.join(user_folder, output_name), index=False)
print("《全部设备耗材流向跟踪表》已输出")