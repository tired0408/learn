"""
知识星球博主的数据抓取
"""
import os
import time
import shutil
import requests
import collections
from tqdm import tqdm
from typing import List
from bs4 import BeautifulSoup
from concurrent.futures import Future
from concurrent.futures import ThreadPoolExecutor
from selenium.webdriver import Chrome
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import NoSuchElementException
from selenium.webdriver.remote.webelement import WebElement


class Store:

    def __init__(self, name):
        dir_path = r"E:\NewFolder\zhishi"
        self.txt_path = os.path.join(dir_path, f"{name}.txt")
        self.img_path = self.init_folder(dir_path, f"{name}的图片")
        self.img_index = 0
        self.img_pool = ThreadPoolExecutor(max_workers=10)
        self.img_tasks = collections.deque()
        self.annex_path = self.init_folder(dir_path, f"{name}的附件")
        self.annex_download_list = collections.deque()
        self.chrome_download = r"D:\Download"
        self.comment_path = self.init_folder(dir_path, f"{name}的评论")
        self.comment_index = 0
        self.f = open(self.txt_path, 'w', encoding='utf-8')

    def __del__(self):
        self.f.close()

    def init_folder(self, folder, name):
        """初始化文件夹"""
        path = os.path.join(folder, name)
        if os.path.exists(path):
            return path
        os.mkdir(path)
        return path

    def write_comment(self, values):
        """写入并生成评论文件夹"""
        self.comment_index += 1
        name = f"{self.comment_index}.txt"
        txt_path = os.path.join(self.comment_path, name)
        with open(txt_path, "w", encoding="utf-8") as f:
            for vl in values:
                f.write(f"{vl}\n")
        return name

    def download_img(self, src):
        """异步下载图片"""
        self.img_index += 1
        img_name = f"{self.img_index}.jpg"
        save_path = os.path.join(self.img_path, img_name)
        if os.path.exists(save_path):
            return img_name
        task = self.img_pool.submit(self.download_img_method, src, save_path)
        self.img_tasks.append(task)
        return img_name

    def download_img_method(self, src, target):
        """下载图片的方法"""
        try:
            response = requests.get(src)
            if response.status_code == 200:
                with open(target, "wb") as f:
                    f.write(response.content)
            else:
                return [src, target]
        except requests.exceptions.ConnectionError:
            return [src, target]

    def wait_start_annex_download(self, name):
        """等待附件开始下载"""
        st = time.time()
        while True:
            if os.path.exists(os.path.join(self.chrome_download, f"{name}.crdownload")):
                break
            if os.path.exists(os.path.join(self.chrome_download, f"{name}")):
                break
            time.sleep(1)
            if (time.time() - st) > 10:
                raise Exception("Waiting download start timeout.")

    def annex_exists(self, name):
        """判断附件是否存在"""
        target = os.path.join(self.annex_path, name)
        if os.path.exists(target):
            return True
        source = os.path.join(self.chrome_download, name)
        if os.path.exists(source):
            shutil.move(source, target)
            return True
        self.annex_download_list.append(name)
        return False

    def wait_img_task(self):
        """等待图片保存线程完成"""
        while len(self.img_tasks) != 0:
            for _ in tqdm(range(len(self.img_tasks))):
                task: Future = self.img_tasks.popleft()
                value = task.result()
                if value is None:
                    continue
                task = self.img_pool.submit(self.download_img_method, *value)
                self.img_tasks.append(task)

    def wait_annex(self):
        """等待附件保存完成"""
        while len(self.annex_download_list) != 0:
            for _ in tqdm(range(len(self.annex_download_list))):
                name = self.annex_download_list.popleft()
                src = os.path.join(self.chrome_download, name)
                if os.path.exists(src):
                    tge = os.path.join(self.annex_path, name)
                    shutil.move(src, tge)
                else:
                    self.annex_download_list.append(name)
                    time.sleep(1)

    def write_info(self, name, date, ctype, content, images, annexs, comment):
        """写入文字信息"""
        self.f.writelines([
            f"**人物:{name}\n",
            f"**时间:{date}\n",
            f"**类型:{ctype}\n",
            "**内容:\n",
            f"{content}\n"
        ])
        if comment is not None:
            self.f.write(f"**评论:{comment}\n")
        if len(images) != 0:
            self.f.write(f"**图片:{','.join(images)}\n")
        if len(annexs) != 0:
            self.f.write(f"**附件:\n{'\n'.join(annexs)}\n")
        self.f.write(f"{'-' * 100}\n")


class Crawler:

    def __init__(self, name, only_owner=False, is_pdf=False, is_img=False, is_comment=False):
        """
        初始化
        :param name: (str); 星球名称
        :param only_owner: (bool); 是否只看星主
        :param is_pdf: (bool); 是否下载PDF
        :param is_img: (bool); 是否下载图片
        :param is_comment: (bool); 是否抓取评论
        """
        self.is_pdf = is_pdf
        self.is_img = is_img
        self.is_comment = is_comment
        self.driver = self.init_chrome()  # 定义chrome浏览器驱动
        self.wait = WebDriverWait(self.driver, 120)  # 定义等待器
        self.owner = Store(name)
        self.member = Store(f"{name}_成员") if not only_owner else None

    def init_chrome(self):
        """定义谷歌浏览器"""
        exe_path = r'E:\NewFolder\chromedriver_mac_arm64_114\chromedriver.exe'
        user_path = r'C:\Users\Administrator\AppData\Local\Google\Chrome\User Data'
        service = Service(exe_path)
        chrome_options = Options()
        chrome_options.add_argument(f'user-data-dir={user_path}')  # 指定用户数据目录
        # chrome_options.add_argument("--disable-extensions")  # 禁用扩展
        # chrome_options.add_argument('--headless')  # 设置无界面模式
        # chrome_options.add_argument('--disable-gpu')  # 禁用 GPU 加速
        chrome_options.add_argument('--log-level=3')  # 关闭日志提示
        prefs = {
            "download.default_directory": r"D:\Download",  # 指定下载目录
        }
        chrome_options.add_experimental_option("prefs", prefs)
        # 启动 Chrome 浏览器
        driver = Chrome(service=service, options=chrome_options)
        return driver

    def __del__(self):
        """关闭谷歌浏览器"""
        self.driver.quit()

    def run(self, url, date=None):
        """
        抓取页面
        :param url: (str); 需要抓取的页面地址
        :param date: (str); 抓取的截至日期,例如:2023.02
        """
        # 打开登入页面
        self.driver.get(r"https://wx.zsxq.com/dweb2/login")
        # 等待手动登录完成
        self.wait.until(EC.visibility_of_element_located((By.CLASS_NAME, "user-container")))
        # 打开需要抓取的页面
        self.driver.get(url)
        self.wait_content_load()
        # 点击只看星主
        if self.member is None:
            menu_container = self.driver.find_element(By.CLASS_NAME, "menu-container")
            menu_container.find_element(By.XPATH, "//div[text()=' 只看星主 ']").click()
            self.wait_content_load()
        # 抓取页面信息
        if date is None:
            stop_year, stop_month = 0, 0
        else:
            stop_year, stop_month = date.split(".")
            stop_year, stop_month = int(stop_year), int(stop_month)
        selector_container = self.driver.find_element(By.TAG_NAME, "app-month-selector")
        if selector_container.is_displayed():
            for year_selector in selector_container.find_elements(By.TAG_NAME, "div")[1:]:
                year_name = self.analysis_content(year_selector)
                if "active" not in year_selector.get_attribute("class"):
                    self.driver.execute_script("arguments[0].click();", year_selector)
                    self.wait_content_load()
                for month_selector in year_selector.find_element(By.XPATH, "..").find_elements(By.TAG_NAME, "li"):
                    month_name = self.analysis_content(month_selector)
                    if "active" not in month_selector.get_attribute("class"):
                        self.driver.execute_script("arguments[0].click();", month_selector)
                        self.wait_content_load()
                    try:
                        self.driver.find_element(By.CLASS_NAME, "no-data")
                    except NoSuchElementException:
                        print(f"保存{year_name}年{month_name}的数据")
                        self.single_page_read()
                    else:
                        print(f"没有{year_name}年{month_name}的数据")
                        break
                    if int(year_name) <= stop_year and int(month_name[:-1]) <= stop_month:
                        print(f"抓取数据，截至到:{date}")
                        break
                else:
                    continue
                break
        else:
            print("保存最近的所有数据")
            self.single_page_read()
        print("等待作者相关图片的保存完成")
        self.owner.wait_img_task()
        print("等待作者相关附件的保存完成")
        self.owner.wait_annex()
        if self.member is not None:
            print("等待成员相关图片的保存完成")
            self.member.wait_img_task()
            print("等待成员相关附件的保存完成")
            self.member.wait_annex()
        print("爬虫抓取完成")

    def single_page_read(self):
        """单页面读取"""
        # 滚动加载
        while True:
            # 模拟滚动加载更多
            self.driver.execute_script('window.scrollTo(0, document.body.scrollHeight)')
            c1 = EC.visibility_of_element_located((By.CLASS_NAME, 'no-more'))
            c2 = EC.visibility_of_element_located((By.TAG_NAME, "app-lottie-loading"))
            self.wait.until(EC.any_of(c1, c2))
            # 无更多数据后退出加载
            try:
                self.driver.find_element(by=By.CLASS_NAME, value="no-more")
            except NoSuchElementException:
                pass
            else:
                break
            # 加载更多卡住
            g_len = self.driver.find_elements(
                By.XPATH, "//app-lottie-loading//*[name()='g' and @style='display: block;']")
            if len(g_len) != 3:
                self.driver.execute_script("window.scrollTo(0, -document.body.scrollHeight / 4);")
                time.sleep(2)
        for topic_element in tqdm(self.driver.find_elements(By.TAG_NAME, "app-topic")):
            role = topic_element.find_element(By.CLASS_NAME, "role")
            if self.member is None:
                store = self.owner
            else:
                role_type = role.get_attribute("class").split(" ")[1]
                store = self.owner if role_type == "owner" else self.member
            role_name = role.text
            date = topic_element.find_element(By.CLASS_NAME, "date").text
            content_container = topic_element.find_element(By.TAG_NAME, "app-talk-content")
            content = content_container.find_element(By.TAG_NAME, "div")
            content_type = content.get_attribute("class")
            content_text = self.analysis_content(content.find_element(By.CLASS_NAME, "content"))
            comments = self.analysis_and_write_comment(topic_element, store)
            images = self.analysis_and_download_imgs(content.find_elements(By.TAG_NAME, "img"), store)
            annexs = self.analysis_and_download_annex(topic_element, store)
            store.write_info(role_name, date, content_type, content_text, images, annexs, comments)

    def wait_content_load(self):
        """等待知识内容加载完成"""
        c1 = EC.visibility_of_element_located((By.CLASS_NAME, 'no-more'))
        c2 = EC.visibility_of_element_located((By.TAG_NAME, "app-lottie-loading"))
        c3 = EC.visibility_of_element_located((By.CLASS_NAME, "no-data"))
        self.wait.until(EC.any_of(c1, c2, c3))

    def analysis_content(self, element: WebElement):
        """解析知识内容"""
        element_html = element.get_attribute('outerHTML')
        soup = BeautifulSoup(element_html, 'lxml')
        return soup.text

    def analysis_and_write_comment(self, topic_element: WebElement, store: Store):
        """分析评论"""
        if not self.is_comment:
            return None
        values = topic_element.find_element(By.CLASS_NAME, "comment-box").find_elements(By.TAG_NAME, "app-comment-item")
        if len(values) == 0:
            return None
        rd = []
        for element in values:
            date = element.find_element(By.CLASS_NAME, "time").text
            comment = element.find_element(By.CLASS_NAME, "text").text
            rd.append(f"({date}){comment}")
        return store.write_comment(rd)

    def analysis_and_download_imgs(self, values: List[WebElement], store: Store):
        """分析并下载图片"""
        if not self.is_img:
            return []
        names = []
        for element in values:
            src_path = element.get_attribute("src")
            name = store.download_img(src_path)
            names.append(name)
        return names

    def analysis_and_download_annex(self, container: WebElement, store: Store):
        """分析并下载附件"""
        if not self.is_pdf:
            return []
        try:
            values = container.find_element(By.TAG_NAME, "app-file-gallery").find_elements(By.CLASS_NAME, "item")
            if len(values) == 0:
                return []
            self.driver.execute_script("arguments[0].scrollIntoView(true);", container)
            names = []
            for element in values:
                name = element.find_element(By.CLASS_NAME, "file-name").text
                names.append(name)
                if store.annex_exists(name):
                    continue
                element.click()
                WebDriverWait(container, 10).until(EC.visibility_of_element_located((By.CLASS_NAME, "download")))
                container.find_element(By.CLASS_NAME, "download").click()
                store.wait_start_annex_download(name)
                self.driver.execute_script("document.elementFromPoint(0, 0).click();")
                WebDriverWait(container, 10).until_not(EC.visibility_of_element_located((By.CLASS_NAME, "download")))
            return names
        except NoSuchElementException:
            return []


if __name__ == "__main__":
    url = r"https://wx.zsxq.com/dweb2/index/group/51288484484424"
    name = r"枪出如龙"
    # 是否只看星主,是否下载PDF,是否下载图片,是否抓取评论
    module = Crawler(name, only_owner=True, is_pdf=False, is_img=True, is_comment=True)
    module.run(url, date=None)
