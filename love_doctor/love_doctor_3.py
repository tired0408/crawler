"""
爱医助医的网站数据分析脚本3
"""
import os
import time
import shutil
import datetime
import pandas as pd
import xlsxwriter
from typing import Dict
from decimal import Decimal, ROUND_HALF_UP
from collections import defaultdict
from utils import read_multi_column
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC



user_folder = os.path.dirname(__file__)
print("读取《设备耗材汇总表》")
input_df = read_multi_column(os.path.join(user_folder, "设备耗材汇总.xlsx"))
"""输出数据格式
-- 区域(dict)
    -- 时间(dict)
        -- 类别(dict)
            -- 家数(int)
"""
print("分析《设备耗材汇总表》")
default_dict = {"总家数": 0, "流向覆盖": 0, "有保养的": 0, "保养率≧80%": 0, "保养率≧80%又有耗材产出": 0, "耗材大于1": 0, "耗材同比增长": 0, "耗材环比增长": 0, "同比环比增长": 0}
output_data: Dict[str, Dict[str, Dict]] = defaultdict(lambda: defaultdict(lambda: default_dict))
now_time = datetime.datetime.now()  
now_year_str = str(now_time.year)
for _, row in input_df.iterrows():
    if now_year_str not in row["时间"]: 
        continue
    time_data = output_data[row["归属区域"]][row["时间"]]
    time_data["总家数"] += 1
    if (row["本区域流向"] > 0 and row["跨区域流向"] >= 0) or (row["本区域流向"] == 0 and row["跨区域流向"] > 0):
        time_data["流向覆盖"] += 1
    maintain_value = float(row["设备保养率"][:-1])
    if maintain_value > 0:
        time_data["有保养的"] += 1
    if maintain_value >= 80:
        time_data["保养率≧80%"] += 1
    if maintain_value >= 80 and (row["本区域流向"] + row["跨区域流向"]) > 0:
        time_data["保养率≧80%又有耗材产出"] += 1
    if row["本区域流向"] + row["跨区域流向"] >= 50:
        time_data["耗材大于1"] += 1
    consume_year_year = float(row["耗材同比"][:-1])
    consume_month_month = float(row["耗材环比"][:-1])
    if consume_year_year > 0:
        time_data["耗材同比增长"] += 1
    if consume_month_month > 0:
        time_data["耗材环比增长"] += 1
    if consume_year_year > 0 and consume_month_month > 0:
        time_data["同比环比增长"] += 1
print("输出《设备耗材家数分析表》")
wb = xlsxwriter.Workbook(os.path.join(user_folder, "设备耗材家数分析表.xlsx"))
for region, data1 in output_data.items():
    ws = wb.add_worksheet(region)
    row_i = 0
    for time_value, data2 in data1.items():
        ws.write(row_i, 0, time_value)
        row_i += 1
        for col_i, value in enumerate(["类别", "家数", "总家数占比", "流向覆盖家数占比"]):
            ws.write(row_i, col_i, value)
        row_i += 1
        for label, value in data2.items():
            ws.write(row_i, 0, label)
            ws.write(row_i, 1, value)
            if label == "总家数":
                row_i += 1
                continue
            ratio_1 = Decimal(value) / Decimal(data2["总家数"]) * 100
            ratio_1 = ratio_1.quantize(Decimal('0.01'), rounding=ROUND_HALF_UP)
            if label == "流向覆盖":
                ws.write(row_i, 2, ratio_1)
                row_i += 1
                continue
            ratio_2 = Decimal(value) / Decimal(data2["流向覆盖"]) * 100
            ratio_2 = ratio_2.quantize(Decimal('0.01'), rounding=ROUND_HALF_UP)
            ws.write(row_i, 2, ratio_1)
            ws.write(row_i, 3, ratio_2)
            row_i += 1
        # 空白行
        row_i += 1
wb.close()
print("《设备耗材家数分析表》已完成")