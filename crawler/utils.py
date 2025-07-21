"""
爬虫抓取的相关通用工具
"""
from selenium.webdriver import Chrome
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options


def init_chrome(chromedriver_path, chrome_path=None, download_path=None, user_path=None,  is_proxy=False, headless=False):
    """初始化浏览器
    
    args:
        chromedriver_path: (str); 浏览器驱动的地址
        download_path: (str); 下载路径
        chrome_path: (str); 浏览器路径
        is_proxy: (bool); 是否使用代理
    """
    service = Service(chromedriver_path)
    options = Options()
    if chrome_path is not None:
        options.binary_location = chrome_path
    if download_path is not None:
        options.add_experimental_option('prefs', {"download.default_directory": download_path})
    if user_path is not None:
        options.add_argument(f'user-data-dir={user_path}')  # 指定用户数据目录
    if is_proxy:
        options.add_argument('--proxy-server=127.0.0.1:8080')
        options.add_argument('ignore-certificate-errors')
    if headless:
        options.add_argument('--headless')
    options.add_argument('--log-level=3')
    driver = Chrome(service=service, options=options)
    return driver