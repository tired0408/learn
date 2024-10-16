"""
爬虫抓取的相关通用工具
"""
import datetime
from selenium.webdriver.common.by import By
from selenium.webdriver.remote.webelement import WebElement
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver import Chrome
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options

def select_date_1(start_date: datetime.datetime, end_date:datetime.datetime, start_label: WebElement, end_label: WebElement, 
                  start_container: WebElement, end_container: WebElement, date_pattern, day_pattern, last_btn: WebElement, next_btn: WebElement):
    """选择日期: 第一下点击选择开始日期,第二下点击选择结束日期
    args:
        start_date: (datetime.datetime); 开始日期
        end_date: (datetime.datetime); 结束日期
        start_label: (WebElement); 开始日期模块的日期显示标签
        end_label: (WebElement); 结束日期模块的日期显示标签
        start_container: (WebElement); 开始日期模块的容器
        end_container: (WebElement); 结束日期模块的容器
        date_pattern: (str); 日期显示标签的表达式
        day_pattern: (list); 具体几号的XPATH表达式
        last_btn: (WebElement); 跳转到上个月的按钮
        next_btn: (WebElement); 跳转到下个月的按钮
    """
    WebDriverWait(start_label, 10).until(lambda ele: ele.text != "")
    for select_date, label, container in zip([start_date, end_date], [start_label, end_label], [start_container, end_container]):
        while True:
            now_date_str = label.text
            now_date = datetime.datetime.strptime(now_date_str, date_pattern)        
            if now_date.year == select_date.year and now_date.month == select_date.month:
                break
            if now_date.year > select_date.year:
                last_btn.click()
            elif now_date.year < select_date.year:
                next_btn.click()
            elif now_date.month > select_date.month:
                last_btn.click()
            elif now_date.month < select_date.month:
                next_btn.click()
            WebDriverWait(label, 10).until(lambda ele: ele.text != now_date_str and ele.text != "")
        container.find_element(By.XPATH, f"{select_date.day}".join(day_pattern)).click()  
    WebDriverWait(start_container, 10).until(lambda ele: not ele.is_displayed())


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