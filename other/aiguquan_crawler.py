"""
爱股圈的数据抓取
"""
import os
import time
import datetime
import pyautogui
from dateutil.relativedelta import relativedelta
from crawler_util import init_chrome
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By
from selenium.webdriver.remote.webelement import WebElement



class Main:

    def __init__(self) -> None:
        self.driver, self.save_path = self.init_driver() 
    
    def __del__(self):
        print("关闭浏览器")
        self.driver.quit()

    def init_driver(self):
        path = r"E:\NewFolder\aiguquan"
        chrome_download = r"D:\Download"
        chrome_path = os.path.join(path, r"..\chromedriver_mac_arm64_114\chrome114\App\Chrome-bin\chrome.exe")
        chromedriver_path = os.path.join(path, r"..\chromedriver_mac_arm64_114\chromedriver.exe")
        user_path = r'C:\Users\Administrator\AppData\Local\Google\Chrome\User Data'
        driver = init_chrome(chromedriver_path, chrome_download, user_path=user_path, chrome_path=chrome_path, is_proxy=False)
        save_path = os.path.join(path, f"{datetime.datetime.now().strftime('%Y%m%d')}.txt")
        return driver, save_path
    
    def run(self):
        """主方法"""
        # 存储数据
        print("进入网站")
        self.driver.get(r"https://www.aiguquan.com/login/login")
        c1 = EC.element_to_be_clickable((By.XPATH, "//a[text()='切换至手机号登录']"))
        c2 = EC.element_to_be_clickable((By.ID, "chatting_room"))
        ele = WebDriverWait(self.driver, 30).until(EC.any_of(c1, c2))
        # 登录
        if ele.tag_name == "a":
            ele.click()
            ele = WebDriverWait(self.driver, 10).until(EC.visibility_of_element_located((By.ID, "log_name")))
            ele.send_keys("17689466627")
            self.driver.find_element(By.ID, "log_pwd").send_keys("CCZCZR")
            ele = WebDriverWait(self.driver, 10).until(EC.element_to_be_clickable((By.ID, "pwd_login_btn")))
            ele.click()
        # 进入聊天室
        handle_len = len(self.driver.window_handles)
        ele = WebDriverWait(self.driver, 30).until(EC.element_to_be_clickable((By.ID, "chatting_room")))
        ele.click()
        WebDriverWait(self.driver, 30).until(lambda d: len(d.window_handles) > handle_len)
        self.driver.switch_to.window(self.driver.window_handles[-1])
        save_datas = []
        # 下载聊天记录
        ele = WebDriverWait(self.driver, 10).until(EC.element_to_be_clickable((By.XPATH, "//button[text()='只看圈主']/ancestor::div[1]")))
        if "bindThis" not in ele.get_attribute("class"):
            ele.click()
        content_ele = WebDriverWait(self.driver, 10).until(EC.visibility_of_element_located((By.XPATH, "//div[@class='h_agq_page_chatting_body']")))
        content_x, content_y = self.get_element_center(content_ele)
        now_date = datetime.datetime.now()
        for i in range(0, 7):
            ele = WebDriverWait(self.driver, 10).until(EC.element_to_be_clickable((By.XPATH, "//div[contains(@class, 'el-date-editor')]/input[@class='el-input__inner']")))
            ele.click()
            date_container = WebDriverWait(self.driver, 10).until(EC.visibility_of_element_located((By.CLASS_NAME, "el-picker-panel__body-wrapper")))
            if i == 0:
                ele = WebDriverWait(self.driver, 10).until(EC.element_to_be_clickable((By.XPATH, ".//button[text()='昨天']")))
                ele.click()
                continue
            select_date = now_date - relativedelta(days=i)
            select_date_str = select_date.strftime("%m-%d")
            month_label = date_container.find_element(By.XPATH, ".//div[@class='el-date-picker__header']/span[2]")
            if str(select_date.month) not in month_label.text:
                ele = ".//div[@class='el-date-picker__header']/button[contains(@class, 'el-icon-arrow-left')]"
                date_container.find_element(By.XPATH, ele).click()
            ele = f".//span[contains(text(), {select_date.day})]/ancestor::td[contains(@class, 'available')]"
            ele = WebDriverWait(date_container, 10).until(EC.element_to_be_clickable((By.XPATH, ele)))
            ele.click()
            ele = f".//li[contains(@class, 'h_agq_page_li')]//span[contains(text(), '{select_date_str}')]"
            ele = WebDriverWait(content_ele, 10).until(EC.visibility_of_element_located((By.XPATH, ele)))
            last_ele_time, same_num = None, 0
            while same_num < 10:
                latest_ele = content_ele.find_elements(By.TAG_NAME, "li")[0]
                if "empty_list" in latest_ele.get_attribute("class"):
                    same_num += 1
                else:
                    latest_ele_time = latest_ele.find_element(By.CLASS_NAME, "h_agq_page_li_time")
                    if latest_ele_time.text == last_ele_time:
                        same_num += 1
                    else:
                        same_num = 0
                        last_ele_time = latest_ele_time.text
                self.driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", latest_ele_time)
                pyautogui.moveTo(content_x, content_y)
                pyautogui.scroll(500)
                time.sleep(0.2)
            datas = content_ele.find_elements(By.TAG_NAME, "li")
            
            for data in datas:
                if "empty_list" in latest_ele.get_attribute("class"):
                    continue
                data_time = data.find_element(By.CLASS_NAME, "h_agq_page_li_time").text
                if select_date_str not in data_time:
                    continue
                data_content = data.find_element(By.CLASS_NAME, "h_agq_page_li_content_box").text
                save_datas.append(f"{data_time}\n{data_content}")
        # 存储数据
        with open(self.save_path, 'a', encoding='utf-8') as f:
            for sd in save_datas:
                f.write(f"{sd}\n")
                f.write(f"{'-' * 100}\n")

    def get_element_center(self, element: WebElement):
        """获取元素中心位置"""
        window_rect = self.driver.get_window_rect()
        window_x = window_rect['x']  # 浏览器窗口的左上角 X 坐标
        window_y = window_rect['y']  # 浏览器窗口的左上角 Y 坐标
        element_location = element.location
        element_size = element.size
        element_center_x = window_x + element_location['x'] + element_size['width'] / 2
        element_center_y = window_y + element_location['y'] + element_size['height'] / 2
        return element_center_x, element_center_y
    

if __name__ == "__main__":
    main_class = Main()
    main_class.run()
