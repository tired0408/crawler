import re
import os
import time
import json
import random
import asyncio
import nodriver as uc
from tqdm import tqdm
from nodriver.core.connection import ProtocolException
from datetime import datetime, timedelta
from nodriver.core.tab import Tab


class XueQiuCrawler:

    def __init__(self):
        # 用户自定义配置
        self.user_path = r'C:\Users\Administrator\AppData\Local\Google\Chrome\User Data'
        self.chrome_path = r"E:\py-workspace\crawler\chromedriver_mac_arm64_114\chrome114\App\Chrome-bin\chrome.exe"
        self.person_id = "2292705444"
        self.url = f'https://xueqiu.com/u/{self.person_id}'
        self.detail_base_url = f"https://xueqiu.com/{self.person_id}"
        self.save_name = "metalslime"
        self.start_date = None
        self.end_date = datetime(year=2025, month=1, day=1)
        # 程序缓存数据
        self.browser = None
        self.page: Tab = None
        self.tmp_data_id_path = "tmp_data_id.json"  # 获取帖子内容的断点
        self.tmp_page_path = "tmp_page.json"  # 获取帖子ID的断点

    async def main(self):
        """主方法"""
        # 1. 启动浏览器 (默认会使用最佳实践配置来降低检测风险)
        #    你可以通过参数控制，例如 headless=True 开启无头模式
        self.browser = await uc.start(
            browser_executable_path=self.chrome_path,
            user_data_dir=self.user_path,
            headless=False
        ) # 建议先设为 False 观察运行情况
        # 获取帖子 ID列表
        success, all_posts = await self.get_all_posts()
        if success:
            # 根据帖子ID获取帖子内容
            await self.get_content(all_posts)
        print("关闭浏览器")
        self.browser.stop()

    async def get_content(self, all_posts: dict):
        """根据POST_ID列表获取帖子内容"""
        already_key = []
        # 保存为TXT数据
        with open(f"{self.save_name}.txt", 'a', encoding='utf-8') as f:
            for time_str, data_id in tqdm(all_posts.items(), desc="博客提取中"):
                if "(" in data_id:
                    f.writelines([
                        f"**时间:{time_str}\n",
                        f"**内容:{data_id}\n",
                        f"{'-' * 100}\n"
                    ])
                    already_key.append(time_str)
                    continue
                each_url = f"{self.detail_base_url}/{data_id}"
                self.page = await self.browser.get(each_url)
                wait_time = random.uniform(1, 5) 
                await asyncio.sleep(wait_time)
                # 等待帖子列表容器出现（根据实际情况调整选择器）
                try:
                    await self.robust_wait_for('.article__bd__detail', timeout=60)
                    detail_ele = await self.page.query_selector(".article__bd__detail")
                    content = detail_ele.text_all
                    if self.judge_content_need(content):
                        f.writelines([
                            f"**时间:{time_str}\n",
                            f"**内容:{content}\n",
                            f"{'-' * 100}\n"
                        ])
                    already_key.append(time_str)
                except Exception as e:
                    print(f"博客内容提取失败: {e}")
                    await self.page.save_screenshot('error.png')
                    print("缓存未提取的帖子ID")
                    for time_str in already_key:
                        all_posts.pop(time_str)
                    print(f"保存断点数据, 剩余 {len(all_posts)} 条")
                    await self.save_data_id(all_posts)
                    return
            if os.path.exists(self.tmp_data_id_path):
                print("删除断点数据")
                os.remove(self.tmp_data_id_path)

    async def get_all_posts(self):
        """获取帖子ID列表"""
        if os.path.exists(self.tmp_data_id_path) and not os.path.exists(self.tmp_page_path):
            print("有DATA_ID的断点数据, 加载断点数据, 不从网站中爬取")
            with open(self.tmp_data_id_path, "r") as f:
                all_posts = json.load(f)
            return True, all_posts
        print("无DATA_ID的断点数据, 从网站中爬取")
        self.page = await self.browser.get(self.url)
        # 等待帖子列表容器出现（根据实际情况调整选择器）
        try:
            await self.robust_wait_for('article.timeline__item', timeout=60)
            print("帖子列表已加载")
        except Exception as e:
            print(f"等待超时: {e}")
            await self.page.save_screenshot('error.png')
            return []
        # 读取缓存
        if os.path.exists(self.tmp_page_path):
            print("有PAGE_INDEX的断点数据, 加载断点数据, 不从头开始爬取")
            self.start_date, page_index = await self.load_page_cache()
            print("跳转断点位置")
            old_id = await self.first_post_id()
            skip_page_input = await self.page.query_selector('a.pagination__next + input')
            await skip_page_input.scroll_into_view()
            await skip_page_input.click()
            await self.submit_by_enter_key(skip_page_input, str(page_index))
            await self.wait_content_load(old_id)
            # 读取缓存的DATA_ID数据
            with open(self.tmp_data_id_path, "r") as f:
                all_posts = json.load(f)
        else:
            print("无PAGE_INDEX的断点数据, 从头开始爬取")
            page_index = 1
            all_posts = {}
        try:
            latest_date = datetime.now()
            # 4. 获取帖子列表
            while True:
                # 获取当前页面上的帖子元素
                post_elements = await self.page.select_all('article.timeline__item')
                post_elements = post_elements[1:] if len(all_posts) == 0 else post_elements
                print(f"当前帖子数: {len(post_elements)}")
                for post in tqdm(post_elements, desc="抓取中"):
                    time_elem = await post.query_selector('.date-and-source')
                    time_value = await self.str_to_datetime(time_elem.text)
                    data_id = await post.query_selector('a[data-id]')
                    data_id = data_id.attrs["data-id"]
                    if self.start_date is not None and time_value > self.start_date:
                        continue
                    if time_value < self.end_date:
                        print("数据时间已超过两年前，结束爬取")
                        break
                    # 通过...判断是否是全部内容,如果是,进一步判断是否含有量化,没有跳过
                    content_description = await post.query_selector('.content--description')
                    content_description = content_description.text_all
                    if "..." in content_description:
                        time_str = time_value.strftime("%Y-%m-%d %H:%M")
                        all_posts[time_str] = data_id
                    elif self.judge_content_need(content_description):
                        time_str = time_value.strftime("%Y-%m-%d %H:%M")
                        all_posts[time_str] = f"({self.detail_base_url}/{data_id}){content_description}"
                    latest_date = time_value
                else:
                    # 获取第一个帖子的 data-id 以便后续比较
                    old_id = await self.first_post_id()
                    # 点击下一页
                    next_button = await self.page.query_selector('a.pagination__next')  # 根据实际情况调整选择器
                    if not next_button:
                        print("没有下一页了，结束爬取")
                        break
                    await next_button.scroll_into_view()
                    # 等待滚动完成（可选）
                    await asyncio.sleep(0.5)
                    # 再点击
                    await next_button.click()
                    # 等待新内容加载（根据实际情况调整等待条件）
                    await self.wait_content_load(old_id)
                    page_index += 1
                    continue
                break
        except Exception as e:
            print(f"博客列表ID提取失败: {e}")
            await self.page.save_screenshot('error.png')
            print("保存读取列表ID的断点数据")
            await self.save_page_cache(latest_date, page_index)
            await self.save_data_id(all_posts)
            return False, all_posts
        if os.path.exists(self.tmp_page_path):
            print("删除读取DATA_ID列表的断点数据")
            os.remove(self.tmp_page_path)
        print("保存博客内容或索引列表")
        await self.save_data_id(all_posts)
        return True, all_posts

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


    async def save_data_id(self, all_posts):
        """保存DATA_ID的断点数据"""
        with open(self.tmp_data_id_path, "w", ) as f:
            json.dump(all_posts, f, indent=2)

    async def save_page_cache(self, start_date:datetime, page_index):
        """保存爬取博客ID列表的断点数据"""
        save_data = [datetime.strftime(start_date, "%Y-%m-%d %H:%M"), page_index]
        with open(self.tmp_page_path, "w", ) as f:
            json.dump(save_data, f, indent=2)
    
    async def load_page_cache(self):
        """加载爬取博客ID列表的断点数据
        
        Returns:
            str: 爬取起始日期
            int: 当前爬取页面
        """
        with open(self.tmp_page_path, "r") as f:
            cache = json.load(f)
        return datetime.strptime(cache[0], "%Y-%m-%d %H:%M"), cache[1]

    async def first_post_id(self):
        """获取第一个帖子的ID"""
        first_post = await self.page.query_selector('article.timeline__item')
        link = await first_post.query_selector('a[data-id]')
        old_id = link.attrs['data-id']
        return old_id

    async def robust_wait_for(self, selector, timeout=60, retry_interval=0.5):
        """
        健壮地等待元素出现，自动处理 ProtocolException。
        """
        start = asyncio.get_event_loop().time()
        while True:
            try:
                # 尝试等待，每次等待较短时间，便于及时重试
                element = await self.page.wait_for(selector, timeout=5)
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
    def judge_content_need(value:str):
        """判断是否是所需的内容"""
        if "量化" in value:
            return True
        if "ai" in value.lower():
            return True
        return False

if __name__ == '__main__':
    # 由于 asyncio.run() 在某些环境下可能有问题，官方推荐使用 loop().run_until_complete()
    crawler = XueQiuCrawler()
    uc.loop().run_until_complete(crawler.main())