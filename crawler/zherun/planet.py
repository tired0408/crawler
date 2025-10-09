"""
抓取知识星球的数据
"""
import os
import sys
root_path = os.path.dirname(os.path.dirname(os.path.dirname(__file__)))
sys.path.append(root_path)
import re
import time
import shutil
import traceback
import requests
import argparse
import datetime
import win32api
import win32con
from bs4 import BeautifulSoup
from dateutil.relativedelta import relativedelta
from typing import List, TextIO, Optional
from crawler.utils import init_chrome
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.action_chains import ActionChains
from selenium.common.exceptions import TimeoutException
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.chrome.webdriver import WebDriver
from selenium.webdriver.remote.webelement import WebElement


class Store:
    """数据保存类"""
    def __init__(self, dir_path, name, is_picture, is_annex, is_segmentation):
        """存储类

        Args:
            dir_path (str): 数据的存储文件夹地址
            name (str): 知识星球星主名称
            is_segmentation (bool): 是否按日期分割
            is_picture (bool): 是否下载图片
            is_annex (bool): 是否下载附件
        """
        self.dir_path = self.init_folder(dir_path, name)
        self.is_segmentation = is_segmentation

        self.f: Optional[TextIO] = open(os.path.join(self.dir_path, f"{name}.txt"), 'a', encoding='utf-8')
        self.img_path = self.init_folder(self.dir_path, f"{name}的图片") if is_picture else None
        self.img_index = len(os.listdir(self.img_path)) if is_picture else 0
        self.annex_path = self.init_folder(self.dir_path, f"{name}的附件") if is_annex else None
        self.html_rename = {}  # HTML文件整理字典，给文件添加日期

    def init_folder(self, folder, name):
        """初始化文件夹"""
        path = os.path.join(folder, name)
        if not os.path.exists(path):
            os.mkdir(path)
            return path
        return path

    def download_img(self, src):
        """下载图片"""
        self.img_index += 1
        img_name = f"{self.img_index}.jpg"
        # 判断数据是否已存在
        save_path = os.path.join(self.img_path, img_name)
        if os.path.exists(save_path):
            return img_name
        # 下载图片
        st = time.time()
        while time.time() - st < 600:
            try:
                response = requests.get(src)
                if response.status_code != 200:
                    continue
                with open(save_path, "wb") as f:
                    f.write(response.content)
                return img_name
            except requests.exceptions.ConnectionError:
                continue
        raise Exception("Download image timeout.")

    def download_annex(self, ele:WebElement, name, save_name):
        """点击开始下载附件
        
        Args:
            ele (WebElement): 附件下载按钮
            name (str): 附件名称
            save_name (str): 附件保存名称
        """
        source = os.path.join(r"D:\Download", name)
        target = os.path.join(self.annex_path, save_name)
        if os.path.exists(source):
            shutil.move(source, target)
            print(f"附件已提前下载好，移动到指定位置:{save_name}")
            return
        for _ in range(5):
            ele.click()
            st = time.time()
            while (time.time() - st) < 120:
                if not os.path.exists(source):
                    time.sleep(1)
                    continue
                target = os.path.join(self.annex_path, save_name)
                shutil.move(source, target)
                return
        raise Exception(f"Waiting download start timeout:{name}")

    def annex_exists(self, save_name):
        """判断附件是否存在"""
        save_path = os.path.join(self.annex_path, save_name)
        if os.path.exists(save_path):
            return True
        return False
    
    def html_tidy(self, old_name, new_name):
        """html的数据整理"""
        old_path: str = os.path.join(self.annex_path, old_name)
        new_path: str = os.path.join(self.annex_path, new_name)
        # 判断是否下载好了
        st = time.time()
        while (time.time() - st) < 120:
            if not os.path.exists(old_path):
                time.sleep(1)
                continue
            time.sleep(1)
            shutil.move(old_path, new_path)
            return
        raise Exception(f"Waiting download finnish timeout:{old_name}")

    def save_wx_html(self, save_name, url):
        """保存微信网页地址"""
        with open(os.path.join(self.annex_path, save_name), 'w', encoding='utf-8') as f:
            f.write(url)

    def write_info(self, name, date: datetime.datetime, ctype, content, images, annexs, comment):
        """写入文字信息
        Args:
            name: (str); 人物名称
            date: (datetime.datetime); 日期,"%Y-%m-%d %H:%M"
            ctype: (str); 类型
            content: (str); 内容
            images: (list); 图片文件名列表
            annexs: (list); 附件文件名列表
            comment: (list); 评论内容列表
        """
        date_str = date.strftime("%Y-%m-%d")
        if self.is_segmentation and date_str not in self.f.name:
            self.f.close()
            txt_path = os.path.join(self.dir_path, f"{date_str}.txt")
            self.f = open(txt_path, "w", encoding="utf-8")

        self.f.writelines([
            f"**人物:{name}\n",
            f"**时间:{date.strftime('%Y-%m-%d %H:%M')}\n",
            f"**类型:{ctype}\n",
            "**内容:\n",
            f"{content}\n",
        ])
        if len(images) != 0:
            images = ','.join(images)
            self.f.write(f"**图片:{images}\n")
        if len(annexs) != 0:
            annexs = ','.join(annexs)
            self.f.write(f"**附件:\n{annexs}\n")
        if len(comment) !=0:
            self.f.write(f"**评论:\n")
            for value in comment:
                self.f.write(f"{value}\n")
        self.f.write(f"{'-' * 100}\n")

class Crawler:
    """网站爬取类"""
    def __init__(self):
        self.is_html = None
        self.is_img = None
        self.annex_name = None
        self.comment_name = None
    
        self.owner: Store = None  # 星主保存类
        self.member: Store = None  # 会员保存类
        self.driver: WebDriver = None  # 浏览器
        self.actions: ActionChains = None
        self.tag2method = {
            "app-talk-content": self.analysis_talk_or_task,
            "app-task-content": self.analysis_talk_or_task,
            "app-answer-content": self.analysis_answer
        }

    def init_parameter(self, forder, name, is_owner, is_img, is_segmentaion, is_html, annex_name, comment_name):
        """初始化参数信息"""
        self.is_html = is_html
        self.is_img = is_img
        self.annex_name = annex_name
        self.comment_name = comment_name

        is_annex = annex_name is not None or is_html
        self.owner = Store(forder, name, is_img, is_annex, is_segmentaion)
        self.member = Store(forder, f"{name}_member", is_img, is_annex, is_segmentaion) if not is_owner else None

    def init_driver(self, *args, **kwargs):
        """初始化chrome浏览器"""
        self.driver = init_chrome(*args, **kwargs)
        if self.comment_name is not None:
            self.actions = ActionChains(self.driver)

    def login(self, url):
        """跳转指定圈主页面"""
        self.driver.get(r"https://wx.zsxq.com/dweb2/login")
        WebDriverWait(self.driver, 120).until(EC.visibility_of_element_located((By.CLASS_NAME, "user-container")))
        self.driver.get(url)
        self.wait_content_load()

    def run(self, start_date:datetime.datetime=None):
        """抓取星球数据内容
        
        Args:
            start_date: (datetime.datetime); 是否只看星主
        """
        # 点击只看星主
        if self.member is None:
            menu_container = self.driver.find_element(By.CLASS_NAME, "menu-container")
            menu_container.find_element(By.XPATH, "//div[text()=' 只看星主 ']").click()
            self.wait_content_load()
        # 只有最近的数据
        selector_container = self.driver.find_element(By.TAG_NAME, "app-month-selector")
        if not selector_container.is_displayed():
            self.single_page_read(start_date)
            print("保存最近的所有数据")
            return
        # 根据开始日期遍历数据
        today = datetime.datetime.now()
        for year in range(start_date.year, today.year+1):
            year_ele = self.driver.find_element(By.XPATH, "//app-month-selector//div[text()='{}']".format(year))
            if "active" not in year_ele.get_attribute("class"):
                year_ele.click()
                self.wait_content_load()
            start_month = start_date.month if year == start_date.year else 1
            end_month = today.month + 1 if year == today.year else 13
            for month in range(start_month, end_month):
                ele = year_ele.find_element(By.XPATH, f"./parent::li//li[text()='{month}月']")
                if "active" not in ele.get_attribute("class"):
                    ele.click()
                    self.wait_content_load()
                if year == start_date.year and month == start_month:
                    self.single_page_read(start_date)
                else:
                    self.single_page_read()

    def single_page_read(self, start_date=None):
        """单页数据读取
        Args:
            start_date: (datetime.datetime); 开始日期
        """
        # 滚动加载出所有数据
        while True:
            # 模拟滚动加载更多
            self.driver.execute_script('window.scrollTo(0, document.body.scrollHeight)')
            c1 = EC.visibility_of_element_located((By.CLASS_NAME, 'no-more'))
            c2 = EC.visibility_of_element_located((By.TAG_NAME, "app-lottie-loading"))
            WebDriverWait(self.driver, 120).until(EC.any_of(c1, c2))
            # 无更多数据后退出加载
            if len(self.driver.find_elements(by=By.CLASS_NAME, value="no-more")) != 0:
                break
            # 加载更多卡住了
            g_len = self.driver.find_elements(By.XPATH, "//app-lottie-loading//*[name()='g' and @style='display: block;']")
            if len(g_len) != 3:
                self.driver.execute_script("window.scrollTo(0, -document.body.scrollHeight / 4);")
                time.sleep(2)
        # 开始抓取数据
        topics = self.driver.find_elements(By.TAG_NAME, "app-topic")
        topics.reverse()
        # 根据开始日期截取数据
        if start_date is not None:
            print("开始日期不为空,对数据进行截断,去除开始日期之前的数据")
            for index in range(len(topics)):
                topic_element = topics[index]
                date = topic_element.find_element(By.CLASS_NAME, "date").text
                now_date = datetime.datetime.strptime(date, "%Y-%m-%d %H:%M")
                if now_date - start_date > datetime.timedelta(0):
                    topics = topics[index:]
                    break
            else:
                raise Exception("开始日期有问题")
        # 遍历数据
        for topic_element in topics:
            date = datetime.datetime.strptime(topic_element.find_element(By.CLASS_NAME, "date").text, "%Y-%m-%d %H:%M")
            # 判断是否是星主
            role = topic_element.find_element(By.CLASS_NAME, "role")
            store = self.owner if self.member is None or role.get_attribute("class").split(" ")[1] != "owner" else self.owner
            # 判断当前主题类型
            role_name = role.text
            condition = (EC.visibility_of_element_located((By.TAG_NAME, name)) for name in self.tag2method.keys())
            content_container = WebDriverWait(topic_element, 10).until(EC.any_of(*condition))
            content_type = content_container.tag_name
            # 读取所需数据
            content_text = self.tag2method[content_type](content_container)
            comments = self.analysis_comment(topic_element) if self.comment_name is not None else []
            images = self.analysis_and_download_imgs(topic_element, store) if self.is_img else []
            annexs = self.analysis_and_download_annex(topic_element, store, date) if self.annex_name is not None else []
            if self.is_html:
                self.analysis_and_download_html(topic_element, store, date)
            # 写入txt文件中
            store.write_info(role_name, date, content_type, content_text, images, annexs, comments)


    def analysis_talk_or_task(self, content: WebElement) -> str:
        """分析自诉以及作业主题的内容"""
        content_text = self.analysis_text(content.find_element(By.CLASS_NAME, "content"))
        return content_text

    def analysis_answer(self, content: WebElement) -> str:
        """分析问答主题的内容"""
        question_text = self.analysis_text(content.find_element(By.CLASS_NAME, "question"))
        answer_text = self.analysis_text(content.find_element(By.CLASS_NAME, "answer"))
        return f"----问题:{question_text}\n----回答:{answer_text}"

    def analysis_text(self, element: WebElement) -> str:
        """分析text的内容"""
        element_html = element.get_attribute('outerHTML')
        soup = BeautifulSoup(element_html, 'lxml')
        return soup.text

    def analysis_comment(self, topic_element: WebElement):
        """分析评论
        
        Returns:
            List[str]: 评论内容列表
        """
        # 判断是否有评论
        values = topic_element.find_elements(By.XPATH, ".//div[contains(@class, 'comment-box')]/app-comment-item")
        if len(values) == 0:
            return []
        # 点开评论详情
        detail_button = topic_element.find_element(By.XPATH, ".//div[text()='查看详情']")
        self.driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", detail_button)
        detail_button.click()
        rd = []
        # 获取评论列表
        topic_detail = WebDriverWait(self.driver, 30).until(EC.element_to_be_clickable(
            (By.XPATH, "//app-topic-detail//div[@class='topic-detail-panel']")))
        while True:
            if len(topic_detail.find_elements(By.TAG_NAME, "app-lottie-loading")) == 0:
                break
            now_len = len(topic_detail.find_elements(By.CLASS_NAME, "comment-container"))
            self.actions.move_to_element(topic_detail).click()
            self.actions.move_to_element(topic_detail).send_keys(Keys.END).perform()
            try:
                WebDriverWait(topic_element, 1).until(lambda ele: len(
                    ele.find_elements(By.CLASS_NAME, "comment-container")) > now_len)
            except TimeoutException:
                continue
            time.sleep(0.1)
        # 遍历评论并判断是否是所需评论
        for comment_container in topic_detail.find_elements(By.CLASS_NAME, "comment-container"):
            comment_item_list = comment_container.find_elements(By.TAG_NAME, "app-comment-item")
            for comment_item in comment_item_list:
                # 评论名字符合
                comment = comment_item.find_element(By.CLASS_NAME, "comment").text
                if comment == self.comment_name:
                    break
                # 回复名字符合
                refers = comment_item.find_elements(By.CLASS_NAME, "refer")
                refer = None if len(refers) == 0 else refers[0].text
                if refer == self.comment_name:
                    break
            else:
                continue
            for comment_item in comment_item_list:
                date = comment_item.find_element(By.CLASS_NAME, "time").text
                content = comment_item.find_element(By.CLASS_NAME, "text").text
                rd.append(f"({date}){content}")
        self.driver.execute_script("document.elementFromPoint(0, 0).click();")
        WebDriverWait(self.driver, 10).until_not(EC.visibility_of_element_located((By.TAG_NAME, "app-topic-detail")))
        return rd

    def analysis_and_download_imgs(self, topic_element: WebElement, store: Store):
        """分析并下载图片
        
        Returns:
            List[str]: 保存的图片名称
        """
        values= topic_element.find_elements(By.TAG_NAME, "img")
        names = []
        for element in values:
            src_path = element.get_attribute("src")
            name = store.download_img(src_path)
            names.append(name)
        return names

    def analysis_and_download_annex(self, container: WebElement, store: Store, date: datetime.datetime):
        """分析并下载附件
        
        Returns:
            List[str]: 保存的附件名称
        """
        values = container.find_elements(By.XPATH, ".//app-file-gallery//div[contains(@class, 'item')]")
        if len(values) == 0:
            return []
        date_str = date.strftime("%Y%m%d_%H%M%S")
        names = []
        for element in values:
            name = element.find_element(By.CLASS_NAME, "file-name").text
            save_name = f"{date_str}_{name}"
            file_type = name.split(".")[-1]
            if self.annex_name != "all" and file_type not in self.annex_name:
                continue
            names.append(name)
            if store.annex_exists(save_name):
                print(f"该附件已下载:{save_name}")
                continue
            self.driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", element)
            ele = WebDriverWait(element, 20).until(EC.element_to_be_clickable(element))
            ele.click()
            try:
                print(f"下载附件:{name}")
                ele = WebDriverWait(container, 3).until(EC.visibility_of_element_located((By.XPATH, ".//div[text()='下载']")))
                store.download_annex(ele, name, save_name)
            except TimeoutException:
                print(f"该附件已开启内容保护,仅支持在App下载:{name}")
            # 关闭弹窗
            ActionChains(self.driver).move_to_element_with_offset(ele, 800, 0).click().perform()
            WebDriverWait(container, 10).until_not(EC.visibility_of_element_located((By.CLASS_NAME, "file-preview-container")))
        return names

    def analysis_and_download_html(self, topic_element: WebElement, store: Store, date:datetime.datetime):
        """分析并下载html"""
        for link_ele in topic_element.find_elements(By.XPATH, ".//div[contains(@class, 'content')]//a"):
            self.driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", link_ele)
            time.sleep(1)
            # 跳过话题性的链接
            if "hashtag" in link_ele.get_attribute("class"):
                continue
            name = link_ele.text
            if "-" in name:
                continue
            url = link_ele.get_attribute("href")
            # 统计局的链接跳过
            if "www.stats.gov.cn" in url:
                continue
            # 微信公众号等的文章直接另存为地址
            if "weixin" in url or "cctv" in url or "pdf" in url:
                if name.startswith("http"):
                    name = name.split("/")[-1]
                store.save_wx_html(f"{date.strftime('%Y%m%d')}_{name}.txt", url)
                continue
            save_name = f"{date.strftime('%Y%m%d')}_{name}.html"
            if store.annex_exists(f"{name}_files"):
                print(f"该附件已下载:{save_name}")
                continue
            time.sleep(1)
            link_ele.click()
            time.sleep(5)
            self.driver.switch_to.window(self.driver.window_handles[-1])
            name = self.driver.title
            if "查找图书" not in name and "知识星球" not in name:
                # 按下ctrl+s
                win32api.keybd_event(0x11, 0, 0, 0)
                win32api.keybd_event(0x53, 0, 0, 0)
                win32api.keybd_event(0x53, 0, win32con.KEYEVENTF_KEYUP, 0)
                win32api.keybd_event(0x11, 0, win32con.KEYEVENTF_KEYUP, 0)
                time.sleep(1)
                # 按下回车
                win32api.keybd_event(0x0D, 0, 0, 0)
                win32api.keybd_event(0x0D, 0, win32con.KEYEVENTF_KEYUP, 0)
                store.html_tidy(f"{name}.html", save_name)
            else:
                print(f"该页面无权限:{name}")
            self.driver.close()
            self.driver.switch_to.window(self.driver.window_handles[-1])
            time.sleep(1)
            
            
    def wait_content_load(self):
        """等待知识内容加载完成"""
        c1 = EC.visibility_of_element_located((By.CLASS_NAME, 'no-more'))
        c2 = EC.visibility_of_element_located((By.TAG_NAME, "app-lottie-loading"))
        c3 = EC.visibility_of_element_located((By.CLASS_NAME, "no-data"))
        WebDriverWait(self.driver, 120).until(EC.any_of(c1, c2, c3))

if __name__ == "__main__":
    parser = argparse.ArgumentParser()
    parser.add_argument("-o", "--owner", action="store_true", help="是否只看星主,默认查看全部")
    parser.add_argument("-i", "--img", action="store_true", help="是否下载图片,默认不下载")
    parser.add_argument("-s", "--segmentaion", action="store_true", help="是否开启按日期分割文件,默认不开启")
    parser.add_argument("-l", "--html", action="store_true", help="是否下载超链接,默认不下载")

    parser.add_argument("-a", "--annex", type=str, default=None, help="下载附件的后缀名,all代表全部,多个用,间隔")
    parser.add_argument("-c", "--comment", type=str, default=None, help="抓取评论的名字")
    parser.add_argument("-d", "--date", type=str, default="1800.01.01_00.00", help="抓取的开始日期,例如:2022.11.30_11.57")
    parser.add_argument("-u", "--url", type=str, default=None, help="知识星球的URL")
    parser.add_argument("-n", "--name", type=str, default=None, help="知识星球的名称")
    opt = parser.parse_args()
    # 测试代码的时候进行修改
    # opt.owner = True
    # opt.img = False
    # opt.segmentaion = True
    # opt.html = True
    # opt.annex = "all"
    # opt.comment = "初善君"
    # opt.date = "2023.01.01_00.00"
    # opt.url = r"https://wx.zsxq.com/group/452241841488"
    # opt.name = "tugou"
    # 验证参数的合规性
    assert opt.url is not None
    assert opt.name is not None
    opt.date = datetime.datetime.strptime(opt.date, "%Y.%m.%d_%H.%M")
    # 定义初始参数
    dir_path = r"E:\NewFolder\zhishi"
    chrome_path = os.path.join(dir_path, r"..\chromedriver_mac_arm64_114\chrome114\App\Chrome-bin\chrome.exe")
    chromedriver_path = os.path.join(dir_path, r"..\chromedriver_mac_arm64_114\chromedriver.exe")
    download_path = r"D:\Download"
    user_path = r'C:\Users\Administrator\AppData\Local\Google\Chrome\User Data'
    # 启动程序
    crawler = Crawler()
    crawler.init_parameter(dir_path, opt.name, opt.owner, opt.img, opt.segmentaion, opt.html, opt.annex, opt.comment)
    crawler.init_driver(chromedriver_path, download_path=download_path, user_path=user_path, chrome_path=chrome_path, is_proxy=False)
    crawler.login(opt.url)
    crawler.run(opt.date)