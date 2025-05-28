"""
爬虫抓取的相关通用工具
"""
import pandas as pd
from selenium.webdriver import Chrome
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options


def init_chrome(chromedriver_path, download_path, user_path=None, chrome_path=None, is_proxy=True):
    """初始化浏览器
    
    args:
        chromedriver_path: (str); 浏览器驱动的地址
        download_path: (str); 下载路径
        chrome_path: (str); 浏览器路径
        is_proxy: (bool); 是否使用代理
    """
    service = Service(chromedriver_path)
    options = Options()
    if user_path is not None:
        options.add_argument(f'user-data-dir={user_path}')  # 指定用户数据目录
    if chrome_path is not None:
        options.binary_location = chrome_path
    if is_proxy:
        options.add_argument('--proxy-server=127.0.0.1:8080')
        options.add_argument('ignore-certificate-errors')
    options.add_argument('--log-level=3')
    options.add_experimental_option('prefs', {
        "download.default_directory": download_path,  # 指定下载目录
    })
    driver = Chrome(service=service, options=options)
    return driver

def read_multi_column(path):
    """读取两级列名的EXCEL表格"""
    df = pd.read_excel(path, header=[0, 1])
    df.columns = [c1 if c2.startswith("Unnamed") else c2 for c1, c2 in df.columns]
    return df



