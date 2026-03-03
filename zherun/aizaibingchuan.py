import re
import os
import time
import json
import asyncio
import nodriver as uc
from tqdm import tqdm
from nodriver.core.connection import ProtocolException
from dateutil.relativedelta import relativedelta
from datetime import datetime, timedelta
from nodriver.core.tab import Tab, cdp

user_path = r'C:\Users\Administrator\AppData\Local\Google\Chrome\User Data'
chrome_path = r"E:\py-workspace\crawler\chromedriver_mac_arm64_114\chrome114\App\Chrome-bin\chrome.exe"
url = 'https://xueqiu.com/u/4104161666'


# 等待第一个帖子的 data-id 发生变化
async def first_post_id(page:Tab):
    first_post = await page.query_selector('article.timeline__item')
    link = await first_post.query_selector('a[data-id]')
    old_id = link.attrs['data-id']
    return old_id

async def str_to_timestamp(value):
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
    result = re.search(r"昨天 (\d{2}):(\d{2})", value)
    if result is not None:
        return (datetime.now() - timedelta(hours=24)).replace(hour=int(result.group(1)), minute=int(result.group(2)))

    raise Exception(f"无法识别的时间格式：{value}")


async def robust_wait_for(page: Tab, selector, timeout=60, retry_interval=0.5):
    """
    健壮地等待元素出现，自动处理 ProtocolException。
    """
    start = asyncio.get_event_loop().time()
    while True:
        try:
            # 尝试等待，每次等待较短时间，便于及时重试
            element = await page.wait_for(selector, timeout=1)
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

async def get_post_id(page: Tab):
    """获取帖子 ID列表"""
    # 等待帖子列表容器出现（根据实际情况调整选择器）
    try:
        await robust_wait_for(page, 'article.timeline__item', timeout=60)
        print("帖子列表已加载")
    except Exception as e:
        print(f"等待超时: {e}")
        await page.save_screenshot('error.png')
        return []
    # 4. 获取帖子列表并提取信息
    end_date = datetime(year=2025, month=1, day=1)
    all_posts = {}
    while True:
        # 获取当前页面上的帖子元素
        post_elements = await page.select_all('article.timeline__item')
        post_elements = post_elements[1:] if len(all_posts) == 0 else post_elements
        print(f"当前帖子数: {len(post_elements)}")
        for post in tqdm(post_elements, desc="抓取中"):
            time_elem = await post.query_selector('.date-and-source')
            time_value = await str_to_timestamp(time_elem.text)
            data_id = await post.query_selector('a[data-id]')
            data_id = data_id.attrs["data-id"]
            if time_value < end_date:
                print("数据时间已超过两年前，结束爬取")
                return all_posts
            time_str = time_value.strftime("%Y-%m-%d %H:%M")
            all_posts[time_str] = data_id
        else:
            # 获取第一个帖子的 data-id 以便后续比较
            old_id = await first_post_id(page)
            # 点击下一页
            next_button = await page.query_selector('a.pagination__next')  # 根据实际情况调整选择器
            if not next_button:
                print("没有下一页了，结束爬取")
                break
            await next_button.scroll_into_view()
            # 等待滚动完成（可选）
            await asyncio.sleep(0.5)
            # 再点击
            await next_button.click()
            # 等待新内容加载（根据实际情况调整等待条件）
            st = time.time()
            while time.time() - st < 60:
                await asyncio.sleep(1)
                new_id = await first_post_id(page)
                if new_id != old_id:
                    print("新内容已加载")
                    old_id = new_id
                    break
            else:
                raise Exception("下一页等待超时")
            continue

async def save_breakpoint(all_posts, tmp_file):
    """保存断点数据"""
    with open(tmp_file, "w", ) as f:
        json.dump(all_posts, f, indent=2)


async def main():
    # 1. 启动浏览器 (默认会使用最佳实践配置来降低检测风险)
    #    你可以通过参数控制，例如 headless=True 开启无头模式
    browser = await uc.start(
        browser_executable_path=chrome_path,
        user_data_dir=user_path,
        headless=False
    ) # 建议先设为 False 观察运行情况

    # 获取帖子 ID列表
    tmp_file = "tmp_data_id.json"
    if os.path.exists(tmp_file):
        print("有断点数据, 加载断点数据, 不从网站中爬取")
        with open(tmp_file, "r") as f:
            all_posts = json.load(f)
    else:
        print("无断点数据, 从网站中爬取")
        # 2. 打开新标签页并访问雪球网目标页面
        page = await browser.get(url)
        all_posts = await get_post_id(page)
    # 保存断点数据
    await save_breakpoint(all_posts, tmp_file)
    already_key = []
    # 保存为TXT数据
    with open(f"aizaibingchuan.txt", 'a', encoding='utf-8') as f:
        for time_str, data_id in tqdm(all_posts.items(), desc="博客提取中"):
            each_url = f"https://xueqiu.com/4104161666/{data_id}"
            page = await browser.get(each_url)
            await asyncio.sleep(5)
            # 等待帖子列表容器出现（根据实际情况调整选择器）
            try:
                await robust_wait_for(page, '.article__bd__detail', timeout=60)
                detail_ele = await page.query_selector(".article__bd__detail")
                f.writelines([
                    f"**时间:{time_str}\n",
                    f"**内容:{detail_ele.text_all}\n",
                ])
                f.write(f"{'-' * 100}\n")
                already_key.append(time_str)
            except Exception as e:
                print(f"博客内容提取失败: {e}")
                await page.save_screenshot('error.png')
                browser.stop()
                print("缓存未提取的帖子ID")
                for time_str in already_key:
                    all_posts.pop(time_str)
                print(f"保存断点数据, 剩余 {len(all_posts)} 条")
                await save_breakpoint(all_posts, tmp_file)
                return
    # 6. 关闭浏览器
    print("关闭浏览器")
    browser.stop()
    print("删除断点数据")
    os.remove(tmp_file)
    print("爬取完成")

if __name__ == '__main__':
    # 由于 asyncio.run() 在某些环境下可能有问题，官方推荐使用 loop().run_until_complete()
    uc.loop().run_until_complete(main())