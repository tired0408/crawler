"""
爱医助医的网站数据分析脚本2
"""
import os
import time
import shutil
import datetime
import xlsxwriter
import pandas as pd
from tqdm import tqdm
from openpyxl import load_workbook
from typing import Dict
from utils import read_multi_column, read_by_openpyxl
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC



user_folder = os.path.dirname(__file__)
print("读取《设备耗材汇总.xlsx》")
wb, ws, header = read_by_openpyxl(os.path.join(user_folder, "设备耗材汇总.xlsx"))
print("分析《设备耗材汇总.xlsx》")
ws_max_row = ws.max_row
output_dict: Dict[str, Dict] = {}
end_time = None
st = time.time()
for row in tqdm(ws.iter_rows(min_row=3, values_only=True), total=ws_max_row-2, desc="分析中"):
    row_dict = dict(zip(header, row))
    sum_data = row_dict["本区域流向"] + row_dict["跨区域流向"]
    key = f'{row_dict["归属区域"]}_{row_dict["省份"]}_{row_dict["客户"]}'
    if key not in output_dict:
        output_dict[key] = {row_dict["时间"]: sum_data}
    else:
        output_dict[key][row_dict["时间"]] = sum_data
    end_time = row_dict["时间"]
wb.close()
end_time = datetime.datetime.strptime(end_time, "%Y-%m")
print("读取《数据库-客户分类目录.xlsx》")
wb, ws, header = read_by_openpyxl(os.path.join(user_folder, "数据库-客户分类目录.xlsx"), header=1)
print("分析《数据库-客户分类目录.xlsx》")
ws_max_row = ws.max_row
name2classify = {}
for row in tqdm(ws.iter_rows(min_row=2, values_only=True), total=ws_max_row-1, desc="分析中"):
    row_dict = dict(zip(header, row))
    name2classify[row_dict["客户名称"]] = row_dict["终端分类"]
print("转换成《全部设备耗材流向跟踪表》所需的数据, 输出到新表格中")
wb = xlsxwriter.Workbook(os.path.join(user_folder, "全部设备耗材流向跟踪表.xlsx"))
ws = wb.add_worksheet()
columns=["归属区域", "省份", "客户"]
now_time = datetime.datetime.now()
for i in range(1, 13):
    columns.append(f"{now_time.year-1}-{i:02d}")
columns.append(f"{now_time.year-1}年合计")
for i in range(1, end_time.month+1):
    columns.append(f"{end_time.year}-{i:02d}")
columns.append(f"{end_time.year}年合计")
columns.append("终端状态")
columns.append("终端分类")
for i, value in enumerate(columns):
    ws.write(0, i, value)
row_i = 0
for key, value in tqdm(output_dict.items(), desc="转换中"):
    each_row = []
    region, province, custom = key.split("_")
    each_row.append(region)
    each_row.append(province)
    each_row.append(custom)
    for i in range(3, 15):
        each_row.append(value.get(columns[i], 0))
    each_row.append(sum(each_row[3:15]))
    end_index = len(columns) - 3
    for i in range(16, end_index):
        each_row.append(value.get(columns[i], 0))
    each_row.append(sum(each_row[16:end_index]))
    # 计算终端状态
    last_year_sum = each_row[columns.index(f"{end_time.year-1}年合计")]
    now_year_sum = each_row[columns.index(f"{end_time.year}年合计")]
    if last_year_sum <= 0 and now_year_sum > 0:
        each_row.append("新增")
    elif last_year_sum > 0 and now_year_sum> 0:
        each_row.append("存量")
    elif last_year_sum > 0 and now_year_sum <= 0:
        each_row.append("丢失")
    else:
        each_row.append("未开发")
    each_row.append(name2classify.get(custom, ""))
    row_i += 1
    for i, value in enumerate(each_row):
        ws.write(row_i, i, value)
wb.close()
print("《全部设备耗材流向跟踪表》已构建完成")