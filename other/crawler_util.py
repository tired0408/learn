"""
爬虫抓取的相关通用工具
"""
import datetime
from selenium.webdriver.common.by import By
from selenium.webdriver.remote.webelement import WebElement
from selenium.webdriver.support.ui import WebDriverWait

def select_date_1(start_date: datetime.datetime, end_date:datetime.datetime, start_label: WebElement, end_label: WebElement, 
                  start_container: WebElement, end_container: WebElement, date_pattern, day_pattern, last_btn: WebElement, next_btn: WebElement):
    """选择日期: 第一下点击选择开始日期,第二下点击选择结束日期"""
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
