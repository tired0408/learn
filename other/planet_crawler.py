"""
知识星球博主的数据抓取
"""
import os
import time
import shutil
import requests
import traceback
import datetime
import collections
from tqdm import tqdm
from typing import List, TextIO, Optional
from bs4 import BeautifulSoup
from datetime import timedelta
from concurrent.futures import Future
from concurrent.futures import ThreadPoolExecutor
from selenium.webdriver import Chrome
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import NoSuchElementException, TimeoutException
from selenium.webdriver.remote.webelement import WebElement
from crawler_util import init_chrome

def init_folder(folder, name):
    """初始化文件夹"""
    path = os.path.join(folder, name)
    if not os.path.exists(path):
        os.mkdir(path)
        return path, 0
    file_num = len(os.listdir(path))
    return path, file_num

class Store:

    def __init__(self, dir_path, name, is_picture, is_annex):
        self.dir_path = dir_path
        self.txt_path = None
        self.f: Optional[TextIO] = None
        self.img_path, self.img_index = init_folder(dir_path, f"{name}的图片") if is_picture else (None, 0)
        self.pool = ThreadPoolExecutor(max_workers=10)
        self.tasks: List[Future] = []
        self.annex_path, _ = init_folder(dir_path, f"{name}的附件") if is_annex else (None, 0)
        self.annex_download_list = collections.deque()
        self.chrome_download = r"D:\Download"

    def __del__(self):
        if self.f is None:
            return
        self.f.close()

    def download_img(self, src):
        """异步下载图片"""
        self.img_index += 1
        img_name = f"{self.img_index}.jpg"
        save_path = os.path.join(self.img_path, img_name)
        if os.path.exists(save_path):
            return img_name
        task = self.pool.submit(self.download_img_method, src, save_path)
        self.tasks.append(task)
        return img_name

    def download_img_method(self, src, target):
        """下载图片的方法"""
        st = time.time()
        while True:
            if time.time() - st > 600:
                raise Exception("Download image timeout.")
            try:
                response = requests.get(src)
                if response.status_code != 200:
                    continue
                with open(target, "wb") as f:
                    f.write(response.content)
                break
            except requests.exceptions.ConnectionError:
                continue

    def wait_start_annex_download(self, name):
        """等待附件开始下载"""
        st = time.time()
        while True:
            if os.path.exists(os.path.join(self.chrome_download, f"{name}.crdownload")):
                break
            if os.path.exists(os.path.join(self.chrome_download, f"{name}")):
                break
            time.sleep(1)
            if (time.time() - st) > 120:
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

    def wait_task(self):
        """等待保存线程完成"""
        for task in self.tasks:
            task.result()

    def wait_annex(self):
        """等待附件保存完成"""
        last_len = len(self.annex_download_list)
        repeat_num = 0
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
            if last_len == len(self.annex_download_list):
                repeat_num += 1
            else:
                last_len = len(self.annex_download_list)
                repeat_num = 0
            if repeat_num > 10:
                print("重复循环多次没有变化，退出循环")
                print(f"问题列表:{self.annex_download_list}")
                break

    def write_info(self, name, date, ctype, content, images, annexs, comment):
        """写入文字信息
        Args:
            name: (str); 人物名称
            date: (str); 日期,"%Y-%m-%d %H:%M"
            ctype: (str); 类型
            content: (str); 内容
            images: (list); 图片文件名列表
            annexs: (list); 附件文件名列表
            comment: (list); 评论内容列表
        """
        day = datetime.datetime.strptime(date, "%Y-%m-%d %H:%M").strftime("%Y-%m-%d")
        txt_path = os.path.join(self.dir_path, f"{day}.txt")
        if txt_path != self.txt_path:
            if self.f is not None:
                self.f.close()
            self.txt_path = txt_path
            self.f = open(txt_path, "w", encoding="utf-8")
        self.f.writelines([
            f"**人物:{name}\n",
            f"**时间:{date}\n",
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

    def __init__(self, name, is_owner=False, is_img=False, annex_name=None, comment_name=None):
        """
        初始化
        :param name: (str); 星球名称
        :param is_owner: (bool); 是否只看星主
        :param is_img: (bool); 是否下载图片
        :param annex_name: (srt); 下载附件的后缀名,用,分割,"all"为全部抓取,None为不抓取
        :param comment_name: (str); 抓取评论的人名
        """
        dir_path = r"E:\NewFolder\zhishi"
        chrome_path = os.path.join(dir_path, r"..\chromedriver_mac_arm64_114\chrome114\App\Chrome-bin\chrome.exe")
        chromedriver_path = os.path.join(dir_path, r"..\chromedriver_mac_arm64_114\chromedriver.exe")
        download_path = r"D:\Download"
        user_path = r'C:\Users\Administrator\AppData\Local\Google\Chrome\User Data'

        self.is_img = is_img
        self.annex_name = annex_name
        self.comment_name = comment_name
        self.driver = init_chrome(chromedriver_path, download_path, user_path=user_path, chrome_path=chrome_path, is_proxy=False)  # 定义chrome浏览器驱动
        self.wait = WebDriverWait(self.driver, 120)  # 定义等待器
        self.actions = ActionChains(self.driver)
        data_folder = os.path.join(dir_path, name)
        if not os.path.exists(data_folder):
            os.mkdir(data_folder)
        self.owner = Store(data_folder, name, is_img, annex_name is not None)
        self.member = Store(data_folder, f"{name}_成员", is_img, annex_name is not None) if not is_owner else None

    def __del__(self):
        """关闭谷歌浏览器"""
        self.driver.quit()

    def run(self, url, date: datetime.datetime = None):
        """
        抓取页面
        :param url: (str); 需要抓取的页面地址
        :param date: (str); 抓取的开始日期,例如:2024.05.10 11:44
        """
        try:
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
            selector_container = self.driver.find_element(By.TAG_NAME, "app-month-selector")
            if selector_container.is_displayed():
                year_ele = selector_container.find_elements(By.TAG_NAME, "div")[1:]
                year_ele.reverse()
                for year_selector in year_ele:
                    year_name = int(self.analysis_text(year_selector))
                    if year_name < date.year:
                        print(f"跳过{year_name}年,开始日期为:{date}")
                        continue
                    if "active" not in year_selector.get_attribute("class"):
                        self.driver.execute_script("arguments[0].click();", year_selector)
                        self.wait_content_load()
                    month_ele = year_selector.find_element(By.XPATH, "..").find_elements(By.TAG_NAME, "li")
                    month_ele.reverse()
                    for month_selector in month_ele:
                        month_name = int(self.analysis_text(month_selector)[:-1])
                        if year_name == date.year and month_name < date.month:
                            print(f"跳过{year_name}年{month_name}月,开始日期为:{date}")
                            continue
                        if "active" not in month_selector.get_attribute("class"):
                            self.driver.execute_script("arguments[0].click();", month_selector)
                            self.wait_content_load()
                        if len(self.driver.find_elements(By.CLASS_NAME, "no-data")) == 0:
                            print(f"保存{year_name}年{month_name}月的数据")
                            if year_name == date.year and month_name == date.month:
                                self.single_page_read(date)
                            else:
                                self.single_page_read()
                        else:
                            print(f"没有{year_name}年{month_name}月的数据")
            else:
                print("保存最近的所有数据")
                self.single_page_read()
        except Exception:
            print(traceback.format_exc())
            print("-" * 150)
            print("读取过程中报错")
            print("-" * 150)
        print("等待作者相关的保存线程完成")
        self.owner.wait_task()
        print("等待作者相关附件的保存完成")
        self.owner.wait_annex()
        if self.member is not None:
            print("等待成员相关的保存线程完成")
            self.member.wait_task()
            print("等待成员相关附件的保存完成")
            self.member.wait_annex()
        print("爬虫抓取完成")

    def single_page_read(self, start_date=None):
        """单页面读取"""
        # 滚动加载
        while True:
            # 模拟滚动加载更多
            self.driver.execute_script('window.scrollTo(0, document.body.scrollHeight)')
            c1 = EC.visibility_of_element_located((By.CLASS_NAME, 'no-more'))
            c2 = EC.visibility_of_element_located((By.TAG_NAME, "app-lottie-loading"))
            self.wait.until(EC.any_of(c1, c2))
            # 无更多数据后退出加载
            if len(self.driver.find_elements(by=By.CLASS_NAME, value="no-more")) != 0:
                break
            # 加载更多卡住
            g_len = self.driver.find_elements(
                By.XPATH, "//app-lottie-loading//*[name()='g' and @style='display: block;']")
            if len(g_len) != 3:
                self.driver.execute_script("window.scrollTo(0, -document.body.scrollHeight / 4);")
                time.sleep(2)
        topics = self.driver.find_elements(By.TAG_NAME, "app-topic")
        topics.reverse()
        if start_date is not None:
            print("开始日期不为空,对数据进行截断")
            for index in tqdm(range(len(topics))):
                topic_element = topics[index]
                date = topic_element.find_element(By.CLASS_NAME, "date").text
                now_date = datetime.datetime.strptime(date, "%Y-%m-%d %H:%M")
                if now_date - start_date > timedelta(0):
                    topics = topics[index:]
                    break
            else:
                raise Exception("开始日期有问题")
        for topic_element in tqdm(topics):
            date = topic_element.find_element(By.CLASS_NAME, "date").text
            role = topic_element.find_element(By.CLASS_NAME, "role")
            if self.member is None:
                store = self.owner
            else:
                role_type = role.get_attribute("class").split(" ")[1]
                store = self.owner if role_type == "owner" else self.member
            role_name = role.text
            for name, text_method in {
                "app-talk-content": self.analysis_talk_or_task,
                "app-task-content": self.analysis_talk_or_task,
                "app-answer-content": self.analysis_answer
            }.items():
                content_container = topic_element.find_elements(By.TAG_NAME, name)
                if len(content_container) == 0:
                    continue
                content_container = content_container[0]
                break
            else:
                raise Exception("该主题内容的格式抓取未定义.")
            content = content_container.find_element(By.TAG_NAME, "div")
            content_type = content.get_attribute("class")
            content_text = text_method(content)
            comments = self.analysis_comment(topic_element)
            images = self.analysis_and_download_imgs(content.find_elements(By.TAG_NAME, "img"), store)
            annexs = self.analysis_and_download_annex(topic_element, store)
            store.write_info(role_name, date, content_type, content_text, images, annexs, comments)

    def wait_content_load(self):
        """等待知识内容加载完成"""
        c1 = EC.visibility_of_element_located((By.CLASS_NAME, 'no-more'))
        c2 = EC.visibility_of_element_located((By.TAG_NAME, "app-lottie-loading"))
        c3 = EC.visibility_of_element_located((By.CLASS_NAME, "no-data"))
        self.wait.until(EC.any_of(c1, c2, c3))

    def analysis_talk_or_task(self, content: WebElement):
        """分析自诉以及作业主题的内容"""
        content_text = self.analysis_text(content.find_element(By.CLASS_NAME, "content"))
        return content_text

    def analysis_answer(self, content: WebElement):
        """分析问答主题的内容"""
        question_text = self.analysis_text(content.find_element(By.CLASS_NAME, "question"))
        answer_text = self.analysis_text(content.find_element(By.CLASS_NAME, "answer"))
        return f"----问题:{question_text}\n----回答:{answer_text}"

    def analysis_text(self, element: WebElement):
        """分析text的内容"""
        element_html = element.get_attribute('outerHTML')
        soup = BeautifulSoup(element_html, 'lxml')
        return soup.text

    def analysis_comment(self, topic_element: WebElement):
        """分析评论"""
        if self.comment_name is None:
            return []
        values = topic_element.find_elements(By.XPATH, ".//div[contains(@class, 'comment-box')]/app-comment-item")
        if len(values) == 0:
            return []
        # TODO 待校验，判断是否有所需评论，没有的话，不点开详情
        if len(topic_element.find_elements(By.XPATH, ".//span[text()='更多评论']")) == 0:
            for comment_item in values:
                if self.judge_comment_need(comment_item):
                    break
            else:
                return []
        # 抓取所需评论
        detail_button = topic_element.find_element(By.XPATH, ".//div[text()='查看详情']")
        self.driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", detail_button)
        detail_button.click()
        rd = []
        pattern = (By.XPATH, "//app-topic-detail//div[@class='topic-detail-panel']")
        WebDriverWait(self.driver, 30).until(EC.element_to_be_clickable(pattern))
        topic_detail = self.driver.find_element(*pattern)
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
        for comment_container in topic_detail.find_elements(By.CLASS_NAME, "comment-container"):
            comment_item_list = comment_container.find_elements(By.TAG_NAME, "app-comment-item")
            for comment_item in comment_item_list:
                if self.judge_comment_need(comment_item):
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

    def judge_comment_need(self, comment_item: WebElement):
        """判断评论是否有需要的内容"""
        comment = comment_item.find_element(By.CLASS_NAME, "comment").text
        refers = comment_item.find_elements(By.CLASS_NAME, "refer")
        refer = None if len(refers) == 0 else refers[0].text
        result = comment == self.comment_name or refer == self.comment_name
        return result

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
        if self.annex_name is None:
            return []
        try:
            values = container.find_element(By.TAG_NAME, "app-file-gallery").find_elements(By.CLASS_NAME, "item")
            if len(values) == 0:
                return []
            names = []
            for element in values:
                name = element.find_element(By.CLASS_NAME, "file-name").text
                file_type = name.split(".")[-1]
                if self.annex_name != "all" and file_type not in self.annex_name:
                    continue
                names.append(name)
                if store.annex_exists(name):
                    continue
                self.driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", element)
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
    import argparse
    parser = argparse.ArgumentParser()
    parser.add_argument("-o", "--owner", action="store_true", help="是否只看星主,默认查看全部")
    parser.add_argument("-i", "--img", action="store_true", help="是否下载图片,默认不下载")
    parser.add_argument("-a", "--annex", type=str, default=None, help="下载附件的后缀名,all代表全部,多个用,间隔")
    parser.add_argument("-c", "--comment", type=str, default=None, help="抓取评论的名字")
    parser.add_argument("-d", "--date", type=str, default="1800.01.01_00.00", help="抓取的开始日期,例如:2022.11.30_11.57")
    parser.add_argument("-u", "--url", type=str, default=None, help="知识星球的URL")
    parser.add_argument("-n", "--name", type=str, default=None, help="知识星球的名称")
    opt = {key: value for key, value in parser.parse_args()._get_kwargs()}
    # 测试代码的时候进行修改
    opt["owner"] = True
    opt["img"] = True
    opt["annex"] = "all"
    opt["comment"] = "司令"
    opt["date"] = "2024.04.01_00.00"
    opt["url"] = r"https://wx.zsxq.com/group/828288122112"
    opt["name"] = "juewushe"
    # 验证参数的合规性
    assert opt["url"] is not None
    assert opt["name"] is not None
    opt["date"] = datetime.datetime.strptime(opt["date"], "%Y.%m.%d_%H.%M")
    module = Crawler(opt["name"], is_owner=opt["owner"], is_img=opt["img"], annex_name=opt["annex"],
                     comment_name=opt["comment"])
    module.run(opt["url"], date=opt["date"])
