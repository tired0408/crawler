import re
import os
import time
import json
import random
import asyncio
import nodriver as uc
from tqdm import tqdm
from nodriver.core.connection import ProtocolException
from nodriver.core.element import Element
from datetime import datetime, timedelta
from nodriver.core.tab import Tab
from nodriver.core.browser import Browser


class FileReader:
    _path: str = "default.txt"

    @classmethod
    def initialize(cls, path):
        cls._path = path

    @classmethod
    def save(cls, time_str, content):
        """保存数据:追加模式"""
        with open(cls._path, 'a', encoding='utf-8') as f:
            f.writelines([f"**时间:{time_str}\n", f"**内容:{content}\n", f"{'-' * 100}\n"])

class TmpData:
    _path = "tmp_data.json"
    page_index = 1
    start_date = None
    end_date = None

    @classmethod
    def load(cls):
        """加载断点数据"""
        if not os.path.exists(cls._path):
            return
        with open(cls._path, "r") as f:
            cache = json.load(f)
            cls.page_index = cache[0]
            cls.start_date = datetime.strptime(cache[1], "%Y-%m-%d %H:%M")
            cls.end_date = datetime.strptime(cache[2], "%Y-%m-%d %H:%M")
            
        
    @classmethod
    def save(cls):
        """缓存断点数据"""
        save_data = [
            cls.page_index,
            datetime.strftime(cls.start_date, "%Y-%m-%d %H:%M"),
            datetime.strftime(cls.end_date, "%Y-%m-%d %H:%M"),
        ]
        with open(cls._path, "w", ) as f:
            json.dump(save_data, f, indent=2)

    @classmethod
    def delete(cls):
        """删除断点数据"""
        if os.path.exists(cls._path):
            os.remove(cls._path)

class XueQiuCrawler:

    def __init__(self):
        # 用户自定义配置
        self.user_path = r'C:\Users\Administrator\AppData\Local\Google\Chrome\User Data'
        self.chrome_path = r"E:\py-workspace\crawler\chromedriver_mac_arm64_114\chrome114\App\Chrome-bin\chrome.exe"
        self.person_id = "3203289965"  # 6157001146, 4185949384
        self.url = f'https://xueqiu.com/u/{self.person_id}'
        self.detail_base_url = f"https://xueqiu.com/{self.person_id}"
        FileReader.initialize("charles_capital.txt")  #zaikan, hexinluoji
        TmpData.start_date = None
        TmpData.end_date = datetime(year=2025, month=1, day=1)
        # 程序缓存数据
        self.browser: Browser = None
        self.page: Tab = None

    async def main(self):
        """主方法"""
        # 1. 启动浏览器 (默认会使用最佳实践配置来降低检测风险)
        #    你可以通过参数控制，例如 headless=True 开启无头模式
        self.browser: Browser = await uc.start(
            browser_executable_path=self.chrome_path,
            user_data_dir=self.user_path,
            headless=False
        )
        print("打开浏览器")
        self.page = await self.browser.get(self.url)
        print("等待加载完成")
        if not (elem:=await self.robust_wait_for('article.timeline__item', timeout=60)):
            return
        print("加载缓存")
        TmpData.load()
        if TmpData.page_index != 1:
            print("跳转指定页面位置")
            old_id = await self.first_post_id()
            skip_page_input = await self.page.query_selector('a.pagination__next + input')
            await skip_page_input.scroll_into_view()
            await skip_page_input.click()
            await self.submit_by_enter_key(skip_page_input, str(TmpData.page_index))
            await self.wait_content_load(old_id)
        print("开始抓取帖子数据")
        try:
            await self.loop_page()
        except Exception as e:
            print(f"博客数据提取失败: {e}")
            await self.page.save_screenshot('error.png')
            print("缓存断点数据")
            TmpData.save()
        else:
            print("数据已提取完毕,删除断点数据")
            TmpData.delete()
        print("关闭浏览器")
        self.browser.stop()

    async def loop_page(self):
        """循环抓取页面数据的方法"""
        while True:
            # 获取当前页面上的帖子元素
            post_elements = await self.page.select_all('article.timeline__item')
            post_elements = post_elements[1:] if TmpData.page_index == 1 else post_elements
            print(f"当前帖子数: {len(post_elements)}")
            for post in tqdm(post_elements, desc=f"抓取第{TmpData.page_index}页"):
                time_elem = await post.query_selector('.date-and-source')
                time_value = await self.str_to_datetime(time_elem.text)
                data_id = await post.query_selector('a[data-id]')
                data_id = data_id.attrs["data-id"]
                if TmpData.start_date is not None and time_value >= TmpData.start_date:
                    continue
                if time_value < TmpData.end_date:
                    print("数据时间已超过两年前，结束爬取")
                    return
                # 通过...判断是否是全部内容,如果是,进一步判断是否含有量化,没有跳过
                content_description_ele = await post.query_selector('.content--description')
                content_description = content_description_ele.text_all
                if "// @" in content_description:
                    content_description = content_description.split("// @")[0]
                # 处理回复的博客内容
                if "..." not in content_description:
                    self.judge_and_save(data_id, time_value, content_description)
                elif "展开 \ue63c" in content_description:
                    content_description = await self.post_expand(post)
                    self.judge_and_save(data_id, time_value, content_description)
                else:
                    content_description = await self.post_new_tab(post)
                    self.judge_and_save(data_id, time_value, content_description)
                TmpData.start_date = time_value
                raise Exception("test")
            print(f"第{TmpData.page_index}页抓取完毕,跳转下一页")
            old_id = await self.first_post_id()
            next_button = await self.page.query_selector('a.pagination__next')
            if not next_button:
                print("没有下一页了，结束爬取")
                return
            await next_button.scroll_into_view()
            # 等待滚动完成（可选）
            await asyncio.sleep(0.5)
            # 再点击
            await next_button.click()
            # 等待新内容加载（根据实际情况调整等待条件）
            await self.wait_content_load(old_id)
            TmpData.page_index += 1
            continue

    async def post_new_tab(self, post: Element):
        """点击进入新标签页的帖子数据"""
        data_id = await post.query_selector('a[data-id]')
        await data_id.scroll_into_view()
        await data_id.click()
        st = time.time()
        while time.time() - st < 10:
            time.sleep(0.5)
            if len(self.browser.tabs) == 2:
                break
        else:
            raise Exception("新标签页打开超时")
        await asyncio.sleep(3)
        new_page = self.browser.tabs[1]
        await self.robust_wait_for_normal(new_page, '.article__bd__detail', timeout=60)
        detail_ele = await new_page.query_selector(".article__bd__detail")
        await new_page.close()
        st = time.time()
        while time.time() - st < 10:
            time.sleep(0.5)
            if len(self.browser.tabs) == 1:
                break
        else:
            raise Exception("新标签页关闭超时")
        await asyncio.sleep(0.5)
        return detail_ele.text_all

    async def post_expand(self, post: Element):
        """含展开的帖子数据抓取"""
        await post.scroll_into_view()
        c1 = '.timeline__expand__control'
        c2 = ".timeline__unfold__control"
        await self.click_element(post, c1, c2)
        content_description_ele = await post.query_selector('.content--detail')
        content_description = content_description_ele.text_all
        return content_description
        
    @staticmethod
    async def click_element(container: Element, click_pattern, result_pattern, timeout=30):
        """点击元素"""
        btn = await container.query_selector(click_pattern)
        await btn.click()
        st = time.time()
        while time.time() - st < timeout:
            time.sleep(0.5)
            result = await container.query_selector(result_pattern)
            if result is None:
                continue
            style = result.attrs["style"]
            if "display: none" in style:
                continue
            return result
        raise Exception(f"点击 '{click_pattern}' 超时 {timeout} 秒")

    async def robust_wait_for(self, selector, timeout=60, retry_interval=0.5):
        """等待元素出现，自动处理 ProtocolException。"""
        result = await self.robust_wait_for_normal(self.page, selector, timeout, retry_interval)
        return result

    @staticmethod
    async def robust_wait_for_normal(page: Tab, selector, timeout=60, retry_interval=0.5):
        try:
            start = asyncio.get_event_loop().time()
            while True:
                try:
                    # 尝试等待，每次等待较短时间，便于及时重试
                    element = await page.wait_for(selector, timeout=5)
                    return element
                except ProtocolException:
                    # 遇到节点失效，稍后重试
                    if asyncio.get_event_loop().time() - start > timeout:
                        raise asyncio.TimeoutError(f"等待 '{selector}' 超时 {timeout} 秒")
                    await asyncio.sleep(retry_interval)
                except asyncio.TimeoutError:
                    # 元素尚未出现，继续轮询
                    if asyncio.get_event_loop().time() - start > timeout:
                        raise
                    await asyncio.sleep(retry_interval)
        except Exception as e:
            print(f"等待 '{selector}' 失败：{e}")
            await page.save_screenshot('error.png')
            return None

    async def first_post_id(self):
        """获取第一个帖子的ID"""
        first_post = await self.page.query_selector('article.timeline__item')
        link = await first_post.query_selector('a[data-id]')
        old_id = link.attrs['data-id']
        return old_id

    @staticmethod
    async def str_to_datetime(value):
        """日期内容转化为日期格式"""
        result = re.search(r"\d{4}-\d{2}-\d{2} \d{2}:\d{2}", value)
        if result is not None:
            return datetime.strptime(result.group(), "%Y-%m-%d %H:%M")
        result = re.search(r"\d{2}-\d{2} \d{2}:\d{2}", value)
        if result is not None:
            return datetime.strptime(f"{datetime.now().year}-{result.group()}", "%Y-%m-%d %H:%M")
        result = re.search(r"(\d+)小时前", value)
        if result is not None:
            return datetime.now() - timedelta(hours=int(result.group(1)))
        result = re.search(r"(\d+)分钟前", value)
        if result is not None:
            return datetime.now() - timedelta(minutes=int(result.group(1)))
        result = re.search(r"昨天 (\d{2}):(\d{2})", value)
        if result is not None:
            return (datetime.now() - timedelta(hours=24)).replace(hour=int(result.group(1)), minute=int(result.group(2)))

        raise Exception(f"无法识别的时间格式：{value}")

    @staticmethod
    async def submit_by_enter_key(elem, text=""):
        """
        通用方案：通过 elem.apply 执行 JS 来输入内容并触发 Enter
        完全不依赖 page.send，避免 CDP 协议版本差异导致的报错
        """
        
        # 构造 JS 代码
        # 注意：如果 text 包含引号或换行，需要做转义处理，这里假设是普通搜索词
        safe_text = text.replace("\\", "\\\\").replace("'", "\\'").replace("\n", "\\n")
        
        js_code = f"""
        function() {{
            // 1. 设置值 (如果需要)
            if ('{safe_text}') {{
                this.value = '{safe_text}';
                // 触发 input 事件，确保 React/Vue 等框架能检测到值变化
                this.dispatchEvent(new Event('input', {{ bubbles: true }}));
            }}
            
            // 2. 聚焦
            this.focus();
            
            // 3. 构造并触发 KeyDown (Enter)
            const eventDown = new KeyboardEvent('keydown', {{
                key: 'Enter',
                code: 'Enter',
                keyCode: 13,
                which: 13,
                bubbles: true,
                cancelable: true
            }});
            this.dispatchEvent(eventDown);
            
            // 4. 构造并触发 KeyUp (Enter)
            const eventUp = new KeyboardEvent('keyup', {{
                key: 'Enter',
                code: 'Enter',
                keyCode: 13,
                which: 13,
                bubbles: true
            }});
            this.dispatchEvent(eventUp);
            
            // 5. 某些网站还需要 change 事件
            this.dispatchEvent(new Event('change', {{ bubbles: true }}));
            
            return true;
        }}
        """
        # 执行 JS (this 指向 elem 本身)
        # apply 方法通常是异步的，需要 await
        await elem.apply(js_code)
        return True

    async def wait_content_load(self, old_id):
        """等待博客列表刷新"""
        st = time.time()
        while time.time() - st < 60:
            await asyncio.sleep(1)
            new_id = await self.first_post_id()
            if new_id != old_id:
                print("新内容已加载")
                old_id = new_id
                break
        else:
            raise Exception("下一页等待超时")

    def judge_and_save(self, data_id, time_value: datetime, value:str):
        """判断是否是所需的内容并保存"""
        if "量化" not in value and "语料" not in value:
            return
        time_str = time_value.strftime("%Y-%m-%d %H:%M")
        content = f"({self.detail_base_url}/{data_id}){value}"
        FileReader.save(time_str, content)

if __name__ == '__main__':
    # 由于 asyncio.run() 在某些环境下可能有问题，官方推荐使用 loop().run_until_complete()
    crawler = XueQiuCrawler()
    uc.loop().run_until_complete(crawler.main())