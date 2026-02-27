"""
爱医助医的网站数据抓取脚本
数据报告-客户维护
"""
import os
import time
import collections
import datetime
import shutil
from calendar import monthrange
from utils import init_chrome
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC


class CrawlerDriver:

    def __init__(self, chromedriver_path, download_path, chrome_path):
        self.download_dir = download_path
        self.driver = init_chrome(chromedriver_path, download_path=download_path, chrome_path=chrome_path, is_proxy=False)
    
    def login(self, acount, password):
        """登录网站"""
        pattern = (By.XPATH, "//button[contains(text(), '登录')]")
        print("打开网页")
        self.driver.get(r"http://cbs.aiyizhuyi.com/ayzy/cbs/yearly_data/monthsum")
        WebDriverWait(self.driver, 10).until(EC.element_to_be_clickable(pattern))
        print("输入账号密码")
        ele = self.driver.find_element(By.NAME, "login")
        ele.clear()
        ele.send_keys(acount)
        ele = self.driver.find_element(By.NAME, "password")
        ele.clear()
        ele.send_keys(password)
        print("登录系统")
        ele = self.driver.find_element(*pattern)
        ele.click()
    
    def click(self, pattern, result, timeout=10, count=3):
        ele = WebDriverWait(self.driver, timeout).until(EC.element_to_be_clickable(pattern))
        for _ in range(count):
            ele.click()
            try:
                if isinstance(result, tuple):
                    ele = WebDriverWait(self.driver, timeout).until(EC.presence_of_element_located(result))
                else:
                    ele = WebDriverWait(self.driver, timeout).until(EC.any_of(*[EC.presence_of_element_located(r) for r in result]))
                return ele
            except:
                continue
        raise Exception(f"点击元素失败, pattern: {pattern}, result: {result}")

    def download(self, pattern, timeout=30):
        """下载数据"""
        download_path = os.path.join(self.download_dir, "客户维护数据报告.xlsx")
        ele = WebDriverWait(self.driver, timeout).until(EC.element_to_be_clickable(pattern))
        for _ in range(3):
            ele.click()
            time.sleep(2)
            st = time.time()
            while time.time() - st < timeout:
                if os.path.exists(download_path):
                    return download_path
                time.sleep(5)
        raise Exception("下载文件失败")
    
def page_custom_table():
    """客户维护表格的标识元素"""
    return (By.XPATH, "//th[contains(@class, 'list-cell-name-service_id')]//a[contains(text(), '本年度未保养结算耗材数量')]")

def main(acount, password, date_range:str):

    # 定义基础数据
    start_date, end_date = date_range.split("-")
    start_date = datetime.datetime.strptime(start_date, "%Y.%m")
    end_date = datetime.datetime.strptime(end_date, "%Y.%m")
    if start_date > end_date:
        raise Exception("时间范围错误,开始时间必须小于结束时间")
    path = os.path.dirname(__file__)
    download_path = os.path.join(path, "custom_file")
    if not os.path.exists(download_path):
        os.makedirs(download_path)
    chromedriver_path = os.path.join(path, r"..\chromedriver_mac_arm64_114\chromedriver.exe")
    chrome_path = os.path.join(path, r"..\chromedriver_mac_arm64_114\chrome114\App\Chrome-bin\chrome.exe")
    # 从网站下载数据
    crawler = CrawlerDriver(chromedriver_path, download_path, chrome_path)
    crawler.login(acount, password)
    print("点击《数据报告》模块")
    c1 = (By.XPATH, "//div[contains(@id, 'layout-canvas')]//ul[contains(@class, 'mainmenu-nav')]//span[contains(text(), '数据报告')]")
    c2 = (By.XPATH, "//span[contains(text(), '客户维护')]")
    crawler.click(c1, c2)
    print("点击《客户维护》标签")
    c1 = c2
    c2 = page_custom_table()
    crawler.click(c1, c2)
    print("开始下载数据")
    regions = collections.deque()
    while start_date <= end_date:
        year = int(start_date.year)
        month = int(start_date.month)
        year_ele = crawler.driver.find_element(By.ID, "select2-year-container")
        if year_ele.text != f"{year}年":
            c1 = (By.ID, "select2-year-container")
            c2 = (By.CLASS_NAME, "select2-container--open")
            crawler.click(c1, c2)
            c1 = (By.XPATH, f"//ul[@id='select2-year-results']//li[contains(text(), '{year}')]")
            c2 = (By.XPATH, f"//span[@id='select2-year-container' and normalize-space()='{year}年']")
            crawler.click(c1, c2)
        month_ele = crawler.driver.find_element(By.ID, "select2-month-container")
        if month_ele.text != f"{month}月":
            c1 = (By.ID, "select2-month-container")
            c2 = (By.CLASS_NAME, "select2-container--open")
            crawler.click(c1, c2)
            c1 = (By.XPATH, f"//ul[@id='select2-month-results']//li[contains(text(), '{month}月')]")
            select_item = crawler.driver.find_element(*c1)
            crawler.driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", select_item)
            c2 = (By.XPATH, f"//span[@id='select2-month-container' and normalize-space()='{month}月']")
            crawler.click(c1, c2)
        time.sleep(2)
        if len(regions) == 0:
            table_ele = crawler.driver.find_element(By.ID, "list-scrollable_listwidget")
            rows = table_ele.find_elements(By.TAG_NAME, "tr")
            for row in rows[1:-1]:
                name = row.find_element(By.TAG_NAME, "td").text
                regions.append(name)
        select_name = regions.popleft()
        is_success = download_file(crawler, year, month, select_name, download_path)
        if not is_success:
            regions.appendleft(select_name)
        if len(regions) == 0:
            print(f"下载{year}年{month}月的数据完成")  
            _, days_in_month = monthrange(year, month)
            start_date = start_date.replace(day=days_in_month) + datetime.timedelta(days=1)
    print("数据已下载完成,关闭浏览器")
    crawler.driver.quit()


def download_file(crawler: CrawlerDriver, year, month, select_name, download_path):
    new_file_name = f"客户维护数据报告_{select_name}{year}年{month}月.xlsx"
    new_file_path = os.path.join(download_path, new_file_name)
    if os.path.exists(new_file_path):
        print(f"已存在, 跳过下载:{new_file_name}")
        return True
    print(f"开始下载{year}年{month}月的{select_name}区域数据")
    try:
        c1 = (By.XPATH, f"//a[normalize-space()='{select_name}']/..")
        region_ele = crawler.driver.find_element(*c1)
        crawler.driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", region_ele)
        c2 = [(By.ID, "download"), (By.CLASS_NAME, "exception-name-block")]
        result_ele = crawler.click(c1, c2)
        if result_ele.tag_name == "div":
            print(f"{select_name}区域页面数据异常, 跳过")
            crawler.driver.back()
            return True
        # 判断点开的数据是否正确
        range_ele = crawler.driver.find_element(By.CLASS_NAME, "header_3")
        hope_text = f"{year}年{month:>02}月"
        if hope_text not in range_ele.text:
            print(f"数据范围错误, 期望包含:{hope_text}, 实际显示:{range_ele.text}")
            return False
        else:
            file_path = crawler.download((By.ID, "download"))
            os.rename(file_path, new_file_path)
    except Exception as e:
        print("数据下载出错")
        print(e)
        return False
    # 返回表格页面
    c1 = (By.XPATH, "//a[contains(text(), '客户维护')]")
    c2 = page_custom_table()
    crawler.click(c1, c2)
    return True

if __name__ == "__main__":
    import argparse
    parser = argparse.ArgumentParser()
    parser.add_argument("-a", "--acount", type=str, help="爱医助医的账号", default="18611756193")
    parser.add_argument("-p", "--password", type=str, help="爱医助医的密码", default="secret")
    parser.add_argument("-d", "--date_range", type=str, help="时间范围, 格式: 2025.01-2025.06", default="2025.11-2026.01")
    opt = parser.parse_args()
    main(opt.acount, opt.password, opt.date_range)