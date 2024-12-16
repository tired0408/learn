"""
医药网站抓取的通用方法
"""
import re
import os
import time
import socket
import warnings
import traceback
import threading
import pandas as pd
from typing import List
from queue import Queue
from datetime import datetime
from selenium.common.exceptions import NoSuchElementException, UnexpectedAlertPresentException, TimeoutException
from selenium.webdriver import Chrome
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.remote.webelement import WebElement
from selenium.webdriver.support.ui import WebDriverWait, Select
from selenium.webdriver.support import expected_conditions as EC


warnings.filterwarnings("ignore", category=UserWarning)

def correct_str(value):
    """修正字符串: 清理无用字符"""
    value = str(value)
    return value.strip()


def clear_and_send(element: WebElement, value):
    """清除input框内的值并输入所需数据"""
    element.clear()
    element.send_keys(value)


def start_socket():
    def sock_method():
        with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as s:
            s.bind(('localhost', 12345))
            s.listen()
            print("服务器启动，等待连接...")
            conn, addr = s.accept()
            with conn:
                print(f"连接自: {addr}")
                while True:
                    data = conn.recv(1024)
                    if not data:
                        break
                    data = data.decode()
                    q.put(data)
                    print(f"接收到数据: {data}")
            print(f"关闭连接:{addr}")
    q = Queue(maxsize=1)
    threading.Thread(target=sock_method).start()
    return q


def analyze_website(path, ignore_names):
    """分析网站数据
    return: list
        1. 需要验证码的网站地址列表
        2. 不需要验证码的网站地址列表
        3. 客户名称列表

    """
    websites = pd.read_excel(path)
    websites_by_code, websites_no_code = [], []
    clients_names = []
    # 网站的验证码情况
    code_condition = {
        WEBURL.spfj: False,
        WEBURL.inca: False,
        WEBURL.ly: False,
        WEBURL.xm_tc: True,
        WEBURL.fj_tc: True,
        WEBURL.sm_tc: True,
        WEBURL.druggc: True
    }
    for _, row in websites.iterrows():
        data = {
            "website_url": correct_str(row.iloc[2]),
            "client_name": correct_str(row.iloc[0]),
            "user": correct_str(row.iloc[3]),
            "password": "" if pd.isna(row.iloc[4]) else correct_str(row.iloc[4]),
        }
        if not pd.isna(row.iloc[1]):
            data["district_name"] = correct_str(row.iloc[1])
        if data["client_name"] in ignore_names:
            print(f"忽视的客户名称:{data['client_name']},{data['user']}")
            continue
        
        clients_names.append(data["client_name"])
        if code_condition[data["website_url"]]:
            websites_by_code.append(data)
        else:
            websites_no_code.append(data)
    return websites_by_code, websites_no_code, clients_names


def get_url_success(driver: Chrome, url, element_type, element_value):
    """登录网站"""
    for _ in range(5):
        driver.get(url)
        if driver.title == "502 Bad Gateway":
            time.sleep(1)
            continue
        try:
            WebDriverWait(driver, 20).until(EC.visibility_of_element_located((element_type, element_value)))
        except TimeoutException:
            raise Exception(f"登录超时,请检查网站是否能访问:{url} ")
        break
    else:
        raise Exception(f"多次登入失败，请检查网站是否能访问:{url}")


class CaptchaSocketServer:

    def __init__(self) -> None:
        self.sock = self.init_sock()
        self.conn = None

    def __del__(self):
        if self.conn is not None:
            print("关闭连接")
            self.conn.close()
        print("关闭套接字")
        self.sock.close()

    def init_sock(self):
        sock = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
        sock.settimeout(60)
        sock.bind(('localhost', 12345))
        sock.listen()
        print("服务已启动,监听IP:localhost, 端口: 12345")
        return sock

    def recv(self):
        if self.conn is None:
            self.conn, addr = self.sock.accept()
            print(f"连接自: {addr}")
        data = self.conn.recv(1024)
        data = data.decode()
        print(f"接收到数据: {data}")
        return data


class SPFJWeb:
    """国控系网站的数据抓取"""
    
    def __init__(self, driver, url) -> None:
        self.url = url
        self.driver: Chrome = driver

    def login(self, user, password, district_name):
        get_url_success(self.driver, self.url, By.ID, "login")
        clear_and_send(self.driver.find_element(By.ID, "user"), user)
        clear_and_send(self.driver.find_element(By.ID, "pwd"), password)
        Select(self.driver.find_element(By.ID, "own")).select_by_visible_text(district_name)
        self.driver.find_element(By.ID, "login").click()
        WebDriverWait(self.driver, 10).until(EC.visibility_of_element_located((By.ID, "butn")))
        print(f"[国控系]{user}用户已登录")

    def purchase_sale_stock(self, start_date=None, end_date=None):
        """
        进销存数据抓取
        :return: [(商品名称, 进货数量, 销售数量, 库存数量), ...]
        """
        def deal_func(elements: List[WebElement]):
            product_name = elements[1].text + elements[2].text
            product_name = product_name.replace(" ", "")
            purchase = int(elements[4].text)
            sales = int(elements[5].text)
            inventory = int(elements[7].text)
            return [product_name, purchase, sales, inventory]
        rd = self.get_table_data(deal_func, "进销存汇总", start_date=start_date, end_date=end_date)
        print(f"[国控系]进销存数据抓取已完成, 日期:{start_date}-{end_date}，共抓取{len(rd)}条数据")
        return rd

    def get_inventory(self, start_date=None):
        """
        获取库存数据
        :return: [(商品名称, 库存数量, 批号), ...]
        """
        def deal_func(elements: List[WebElement]):
            product_name = elements[1].text + elements[3].text
            product_name = product_name.replace(" ", "")
            amount = int(elements[4].text)
            code = str(elements[6].text)
            return [product_name, amount, code]
        rd = self.get_table_data(deal_func, "库存数据", start_date)
        print(f"[国控系]库存数据抓取已完成, 日期:{start_date}，共抓取{len(rd)}条数据")
        return rd

    def get_table_data(self, deal_func, table_type, start_date=None, end_date=None):
        """获取表格数据的通用方法"""
        Select(self.driver.find_element(By.ID, "type")).select_by_visible_text(table_type)
        if start_date is not None:
            element = self.driver.find_element(By.ID, "txtBeginDate")
            self.driver.execute_script("arguments[0].value = arguments[1]", element, start_date)
        if end_date is not None:
            element = self.driver.find_element(By.ID, "txtEndDate")
            self.driver.execute_script("arguments[0].value = arguments[1]", element, end_date)
        self.driver.find_element(By.ID, "butn").click()
        WebDriverWait(self.driver, 60).until_not(EC.visibility_of_element_located((By.ID, "loading")))
        rd = []
        try:
            values = self.driver.find_element(By.ID, "customers")
            values = values.find_elements(By.TAG_NAME, "tr")
            for each_v in values[1:]:
                each_v = each_v.find_elements(By.TAG_NAME, "td")
                rd.append(deal_func(each_v))
        except NoSuchElementException:
            pass
        return rd


class INCAWeb:
    """片仔癀漳州医药有限公司的数据抓取"""

    def __init__(self, driver, download_path, url) -> None:
        self.url = url
        self.path = download_path
        self.driver: Chrome = driver

    def login(self, user, password):
        get_url_success(self.driver, self.url, By.ID, "login_link")
        clear_and_send(self.driver.find_element(By.ID, "userName"), user)
        if password == "":
            self.driver.find_element(By.ID, "passWord").clear()
        else:
            clear_and_send(self.driver.find_element(By.ID, "passWord"), password)
        clear_and_send(self.driver.find_element(By.ID, "inputCode"), self.driver.find_element(By.ID, "checkCode").text)
        WebDriverWait(self.driver, 10).until(EC.visibility_of_element_located((By.TAG_NAME, "option")))
        self.driver.find_element(By.ID, "login_link").click()
        WebDriverWait(self.driver, 30).until(EC.visibility_of_element_located((By.ID, "tree1")))
        print(f"[片仔癀漳州]{user}用户已登录")

    def get_inventory(self):
        """
        获取库存数据
        return: [(商品名称, 库存数量, 批号), ...]
        """
        def deal_func(elements: List[WebElement]):
            product_name = elements[1].text + elements[2].text
            product_name = product_name.replace(" ", "")
            amount = int(elements[5].text)
            code = str(elements[6].text)
            return [product_name, amount, code]
        rd = self.get_table_data(deal_func, "供应商网络服务", "库存明细查询")
        print(f"[片仔癀漳州]库存数据抓取已完成，共抓取{len(rd)}条数据")
        return rd

    def get_purchase(self, start_date=None):
        """
        获取进货数据
        return: [(商品名称, 进货数量), ...]
        """
        def deal_func(elements: List[WebElement]):
            product_name = elements[2].text + elements[3].text
            product_name = product_name.replace(" ", "")
            amount = int(elements[9].text)
            return [product_name, amount]
        rd = self.get_table_data(deal_func, "供应商网络服务", "进货明细查询", start_date)
        print(f"[片仔癀漳州]进货数据抓取已完成，共抓取{len(rd)}条数据")
        return rd

    def get_sales(self, start_date=None, end_date=None):
        """
        获取发货数据
        return: [(商品名称, 发货数量), ...]
        """
        def sale_download():
            button = WebDriverWait(self.driver, 60).until(EC.element_to_be_clickable((By.ID, "Button1")))
            button.click()

        self.show_table_data("客户网络服务", "发货明细查询", start_date=start_date, end_date=end_date)
        file_name = wait_download(self.path, "发货明细下载", sale_download)
        time.sleep(1)
        file_path = os.path.join(self.path, file_name)
        data = pd.read_excel(file_path, header=0)
        rd = []
        for _, row in data.iterrows():
            product_name: str = row["商品名称"] + row["规格"]
            product_name = product_name.replace(" ", "")
            amount = int(row["数量"])
            rd.append([product_name, amount])
        print(f"[片仔癀漳州]销售数据抓取已完成, 日期:{start_date}-{end_date}，共抓取{len(rd)}条数据")
        os.remove(file_path)
        return rd

    def get_table_data(self, deal_func, tree_type, table_type, start_date=None, end_date=None):
        """获取表格数据
        :param deal_func: (function); 表格的处理方法
        :param tree_type: (str); 树形菜单栏的名称
        :param table_type: (str); 表格的名称
        :param start_date: (str); 开始日期,格式:"%Y-%m-%d"
        :param end_date: (str); 结束日期
        """
        self.show_table_data(tree_type, table_type, start_date=start_date, end_date=end_date)
        rd = []
        try:
            while True:
                content = self.driver.find_element(By.CLASS_NAME, "bill_m")
                tabel = content.find_elements(By.CLASS_NAME, 'formsT_No_table')[1]
                values = tabel.find_elements(By.TAG_NAME, "tr")
                for each_v in values[1:]:
                    each_v = each_v.find_elements(By.TAG_NAME, "td")
                    rd.append(deal_func(each_v))
                pages = content.find_element(By.CLASS_NAME, "pages")
                page_info = pages.find_element(By.CLASS_NAME, "page_m").text
                index, total_index = page_info.split("/")
                if index == total_index:
                    break
                next_page = pages.find_element(By.CLASS_NAME, "next")
                next_page.click()
        except NoSuchElementException:
            pass
        return rd

    def show_table_data(self, tree_type, table_type, start_date=None, end_date=None):
        """显示表格数据"""
        self.driver.switch_to.default_content()
        try:
            xpath = f"//span[text()='{tree_type}']/preceding-sibling::div[contains(@class, 'l-expandable-close')]"
            self.driver.find_element(By.XPATH, xpath).click()
        except NoSuchElementException:
            print(f"[片仔癀漳州]{tree_type},菜单栏已打开")
        self.driver.find_element(By.XPATH, f"//span[text()='{table_type}']").click()
        WebDriverWait(self.driver, 30).until(EC.visibility_of_element_located(
            (By.XPATH, f"//a[text()='{table_type}']/parent::*")))
        # 跳转到响应的iframe
        iframe_id = self.driver.find_element(By.XPATH, f"//a[text()='{table_type}']/parent::*")
        iframe_id = iframe_id.get_attribute("tabid")
        self.driver.switch_to.frame(self.driver.find_element(By.ID, iframe_id))
        if start_date is None and end_date is None:
            self.driver.find_element(By.ID, "submit_s").click()
        else:
            self.driver.find_element(By.ID, "but_b").click()
            ele = WebDriverWait(self.driver, 30).until(EC.element_to_be_clickable((By.ID, "but_con")))
            ele.click()
            WebDriverWait(self.driver, 30).until(EC.visibility_of_element_located((By.CLASS_NAME, "modal-content")))
            btn = self.driver.find_element(By.XPATH, "//input[@value='shijian']")
            if not btn.is_selected():
                btn.click()
            start_element = self.driver.find_element(By.NAME, "startdate")
            self.driver.execute_script("arguments[0].value = arguments[1]", start_element, start_date)
            end_element = self.driver.find_element(By.NAME, "enddate")
            if end_date is None:
                end_date = datetime.now().strftime("%Y-%m-%d")
            self.driver.execute_script("arguments[0].value = arguments[1]", end_element, end_date)
            self.driver.find_element(By.XPATH, "//div[contains(@class, 'modal-footer')]/input[@value='查询']").click()
        WebDriverWait(self.driver, 30).until_not(EC.visibility_of_element_located((By.CLASS_NAME, "l-tab-loading")))


class LYWeb:
    """鹭燕网站的数据抓取"""

    def __init__(self, driver, url) -> None:
        self.url = url
        self.driver: Chrome = driver

    def login(self, user, password, district_name):
        get_url_success(self.driver, self.url, By.CLASS_NAME, "buttonsubmit")
        ele = WebDriverWait(self.driver, 10).until(EC.element_to_be_clickable((By.NAME, "username")))
        clear_and_send(ele, user)
        ele = WebDriverWait(self.driver, 10).until(EC.element_to_be_clickable((By.NAME, "loginpwd")))
        clear_and_send(ele, password)
        Select(self.driver.find_element(By.NAME, "select")).select_by_visible_text(district_name)
        ele = EC.element_to_be_clickable((By.CLASS_NAME, "buttonsubmit"))
        ele = WebDriverWait(self.driver, 60).until(ele)
        ele.click()
        c1 = EC.presence_of_element_located((By.NAME, "menu"))
        c2 = EC.visibility_of_element_located((By.ID, "proceed-button"))
        ele = WebDriverWait(self.driver, 30).until(EC.any_of(c1, c2))
        if ele.get_attribute("id") == "proceed-button":
            ele.click()
            WebDriverWait(self.driver, 30).until(c1)
        print(f"[鹭燕]{user}用户已登录")

    def get_inventory(self):
        """
        获取库存明细信息数据（该地方的库存总数相同产品不可叠加，）
        :return: [(商品名称, 库存数量, 批号), ...]
        """
        def deal_func(elements: List[WebElement]):
            product_name = elements[0].text + elements[1].text
            product_name = product_name.replace(" ", "")
            amount = int(float(elements[2].text))
            code = str(elements[5].text)
            return [product_name, amount, code]
        rd = self.get_table_data(deal_func, "库存明细信息")
        print(f"[鹭燕]库存数据抓取已完成，共抓取{len(rd)}条数据")
        return rd

    def purchase_sale_stock(self, start_date=None, end_date=None):
        """
        进销存汇总数据抓取
        :return: [(商品名称, 进货数量, 销售数量, 库存数量), ...]
        """
        def deal_func(elements: List[WebElement]):
            product_name = elements[0].text + elements[1].text
            product_name = product_name.replace(" ", "")
            purchase = int(float(elements[4].text))
            sales = int(float(elements[3].text))
            inventory = int(float(elements[5].text))
            return [product_name, purchase, sales, inventory]
        rd = self.get_table_data(deal_func, "进销存汇总表", start_date, end_date)
        print(f"[鹭燕]进销存数据抓取已完成,日期:{start_date}-{end_date}，共抓取{len(rd)}条数据")
        return rd

    def get_table_data(self, deal_func, table_type, start_date=None, end_date=None):
        """获取表格数据"""
        self.driver.switch_to.default_content()
        self.driver.switch_to.frame(self.driver.find_element(By.NAME, "menu"))
        self.driver.find_element(By.XPATH, f"//strong[text()='{table_type}']").click()
        self.driver.switch_to.default_content()
        self.driver.switch_to.frame(self.driver.find_element(By.NAME, "main"))
        if start_date is not None:
            element = self.driver.find_element(By.NAME, "StartDate")
            self.driver.execute_script("arguments[0].value = arguments[1]", element, start_date)
        if end_date is not None:
            element = self.driver.find_element(By.NAME, "EndDate")
            self.driver.execute_script("arguments[0].value = arguments[1]", element, end_date)
        try:
            self.driver.find_element(By.XPATH, "//input[@value='开始查询']").click()
        except NoSuchElementException:
            pass
        rd = []
        try:
            values = self.driver.find_element(By.CLASS_NAME, "gridBody").find_elements(By.TAG_NAME, "tr")
            for each_v in values[1:]:
                each_v = each_v.find_elements(By.TAG_NAME, "td")
                rd.append(deal_func(each_v))
        except NoSuchElementException:
            pass
        return rd


class TCWeb:
    """同春医药网站的数据抓取"""

    def __init__(self, driver, captcha, url) -> None:
        self.url = url
        self.driver: Chrome = driver
        self.captcha: CaptchaSocketServer = captcha

    def login(self, user, password):
        get_url_success(self.driver, self.url, By.ID, "imgLogin")
        while True:
            try:
                WebDriverWait(self.driver, 10).until(EC.visibility_of_element_located((By.ID, "imgLogin")))
                clear_and_send(self.driver.find_element(By.ID, "txt_UserName"), user)
                time.sleep(0.2)
                clear_and_send(self.driver.find_element(By.ID, "txtPassWord"), password)
                time.sleep(0.2)
                captcha_value = self.captcha.recv()
                print(f"输入验证码:{captcha_value}")
                clear_and_send(self.driver.find_element(By.ID, "txtVerifyCode"), captcha_value)
                time.sleep(0.2)
                self.driver.find_element(By.ID, "imgLogin").click()
                WebDriverWait(self.driver, 10).until(EC.presence_of_element_located((By.ID, "btnSearch")))
            except UnexpectedAlertPresentException as e:
                if e.alert_text == "验证码输入出错":
                    continue
            except Exception:
                page = self.driver.page_source
                print(page)
            break
        WebDriverWait(self.driver, 10).until(EC.visibility_of_element_located((By.ID, "btnSearch")))
        print(f"[同春医药]{user}用户已登录")

    def get_inventory(self):
        """
        获取库存数据
        :return: [(商品名称, 库存数量), ...]
        """
        self.driver.find_element(By.XPATH, "//img[@src='images/menu001.gif']").click()
        btn = WebDriverWait(self.driver, 10).until(EC.visibility_of_element_located((By.NAME, "lnkExcel")))
        btn.click()
        WebDriverWait(self.driver, 10).until(EC.visibility_of_element_located((By.XPATH, "//table[@border='1']")))
        rd = []
        try:
            values = self.driver.find_element(By.XPATH, "//table[@border='1']").find_elements(By.TAG_NAME, "tr")
            for each_v in values[1:]:
                each_v = each_v.find_elements(By.TAG_NAME, "td")
                product_name = each_v[1].text + each_v[2].text
                product_name = product_name.replace(" ", "")
                inventory = int(each_v[6].text)
                rd.append([product_name, inventory])
        except NoSuchElementException:
            print("[同春医药]无库存数据")
        print(f"[同春医药]库存数据抓取已完成，共抓取{len(rd)}条数据")
        return rd

    def get_product_flow(self, start_date, end_date):
        """
        获取商品流向数据
        :return: [(商品名称, 进货数量, 销售数量), ...]
        """
        self.driver.find_element(By.XPATH, "//img[@src='images/menu001.gif']").click()
        element = WebDriverWait(self.driver, 60).until(EC.visibility_of_element_located((By.NAME, "time1")))
        self.driver.execute_script("arguments[0].value = arguments[1]", element, start_date)
        element = WebDriverWait(self.driver, 60).until(EC.visibility_of_element_located((By.NAME, "time2")))
        self.driver.execute_script("arguments[0].value = arguments[1]", element, end_date)
        btn = WebDriverWait(self.driver, 60).until(EC.element_to_be_clickable((By.NAME, "btnSearch")))
        btn.click()
        condition = EC.visibility_of_element_located((By.XPATH, "//strong[text()='商品流向列表']"))
        element = WebDriverWait(self.driver, 60).until(condition)
        table = element.find_element(By.XPATH, "./ancestor::*[4]/following-sibling::table")
        rd = []
        try:
            values = table.find_elements(By.TAG_NAME, "tr")
            for tr_ele in values[1:]:
                td_list = tr_ele.find_elements(By.TAG_NAME, "td")
                if td_list[0].text == "":
                    break
                product_name = td_list[3].text + td_list[4].text
                sales = td_list[7].text
                sales = sales.strip()
                sales = 0 if sales == "" else int(sales)
                rd.append([product_name, sales])
        except NoSuchElementException:
            print("[同春医药]无流向数据")
        print(f"[同春医药]流向数据抓取已完成,日期:{start_date}-{end_date}，共抓取{len(rd)}条数据")
        return rd


class DruggcWeb:
    """片仔癀宏仁医药有限公司网站的数据抓取"""

    def __init__(self, driver, captcha, download_path, url) -> None:
        self.url = url
        self.path = download_path
        self.driver: Chrome = driver
        self.captcha: CaptchaSocketServer = captcha

    def login(self, user, password, district_name):
        get_url_success(self.driver, self.url, By.ID, "login")
        clear_and_send(self.driver.find_element(By.ID, "username"), user)
        clear_and_send(self.driver.find_element(By.ID, "password"), password)
        Select(self.driver.find_element(By.ID, "entryid")).select_by_visible_text(district_name)
        while True:
            try:
                captcha = self.captcha.recv()
                print(f"输入验证码:{captcha}")
                if len(captcha) != 4:
                    self.driver.find_element(By.ID, "captchaImg").click()
                    time.sleep(0.1)
                    continue
                clear_and_send(self.driver.find_element(By.ID, "captcha"), captcha)
                button = WebDriverWait(self.driver, 60).until(EC.element_to_be_clickable((By.ID, "login")))
                button.click()
                c1 = EC.visibility_of_element_located((By.XPATH, "//div[text()='验证码不正确！']"))
                c2 = EC.visibility_of_element_located((By.ID, "side-menu"))
                ele = WebDriverWait(self.driver, 5).until(EC.any_of(c1, c2))
                print(f"[片仔癀宏仁医药]等待的元素ID:{ele.get_attribute('id')}")
                if ele.get_attribute("id") == "side-menu":
                    break
                print("[片仔癀宏仁医药]验证码识别错误，更换验证码图片")
                self.driver.find_element(By.ID, "captchaImg").click()
                WebDriverWait(self.driver, 60).until_not(c1)
            except Exception:
                print(traceback.format_exc())
                print(1111111111111111111111111111111)
        WebDriverWait(self.driver, 10).until(EC.visibility_of_element_located((By.ID, "side-menu")))
        print(f"[片仔癀宏仁医药]{user}用户已登录")

    def get_inventory(self):
        """
        获取库存数据
        :return: [(商品名称, 库存数量, 批号), ...]
        """
        def deal_inventory(row):
            product_name: str = row["通用名"] + row["规格"]
            product_name = product_name.replace(" ", "")
            amount = int(row["数量"])
            code = str(row["批号"])
            return [product_name, amount, code]
        self.__search_data("库存明细")
        rd = self.__get_data_by_excel(deal_inventory, "库存明细")
        print(f"[片仔癀宏仁医药]库存数据抓取已完成，共抓取{len(rd)}条数据")
        return rd

    def get_purchase(self, start_date):
        """
        获取进货数据
        :return: [(商品名称, 进货数量), ...]
        """
        def deal_restock(elements: List[WebElement]):
            product_name = elements[2].text + elements[4].text
            product_name = product_name.replace(" ", "")
            amount = int(elements[9].text)
            return [product_name, amount]
        self.__search_data("进货明细", start_date)
        rd = self.__get_data_by_excel(deal_restock, "进货明细")
        print(f"[片仔癀宏仁医药]进货数据抓取已完成, 日期:{start_date}，共抓取{len(rd)}条数据")
        return rd

    def get_sales(self, start_date, end_date=None):
        """
        获取销售数据
        :return: [(商品名称, 销售数量), ...]
        """
        def deal_sales(row):
            product_name: str = row["通用名"] + row["规格"]
            product_name = product_name.replace(" ", "")
            amount = int(row["流向数量"])
            return [product_name, amount]
        self.__search_data('供应商流向', start_date, end_date)
        rd = self.__get_data_by_excel(deal_sales, "供应商流向")
        return rd

    def __search_data(self, table_type, start_date=None, end_date=None):
        """根据条件查询信息"""
        self.driver.switch_to.default_content()
        self.driver.find_element(By.XPATH, f"//span[text()='{table_type}']").click()
        self.driver.switch_to.frame(self.driver.find_element(By.ID, "mainframe"))
        c1 = EC.visibility_of_element_located((By.CLASS_NAME, "fixed-table-loading"))
        WebDriverWait(self.driver, 60).until_not(c1)
        if start_date is not None:
            element = self.driver.find_element(By.ID, "beginCreateTime")
            self.driver.execute_script("arguments[0].value = arguments[1]", element, start_date)
        if end_date is not None:
            element = self.driver.find_element(By.ID, "endCreateTime")
            self.driver.execute_script("arguments[0].value = arguments[1]", element, end_date)
        if start_date is not None or end_date is not None:
            ele = WebDriverWait(self.driver, 5).until(EC.element_to_be_clickable((By.XPATH, "//button[text()='查询']")))
            ele.click()
        condition = EC.visibility_of_element_located((By.CLASS_NAME, "fixed-table-loading"))
        WebDriverWait(self.driver, 60).until_not(condition)

    def get_data_by_table(self, deal_func):
        """从网站上的表格获取数据"""
        rd = []
        while True:
            condition = EC.visibility_of_element_located((By.CLASS_NAME, "fixed-table-loading"))
            WebDriverWait(self.driver, 60).until_not(condition)
            values = self.driver.find_element(By.ID, "dataTable")
            values = values.find_elements(By.TAG_NAME, "tr")
            for each_v in values[1:]:
                each_v = each_v.find_elements(By.TAG_NAME, "td")
                if each_v[0].text == '没有找到匹配的记录':
                    return rd
                rd.append(deal_func(each_v))
            page_info = self.driver.find_element(By.CLASS_NAME, "fixed-table-pagination")
            page_numbers = re.findall(r'\d+', page_info.find_element(By.CLASS_NAME, "pagination-info").text)
            if page_numbers[-1] == page_numbers[-2]:
                break
            page_info.find_element(By.XPATH, "//li[contains(@class, 'page-next')]/a").click()
        return rd

    def __get_data_by_excel(self, deal_func, name):
        """通过导出文件获取数据"""
        def data_download():
            button = WebDriverWait(self.driver, 10).until(EC.element_to_be_clickable((By.ID, "export")))
            button.click()
        file_name = wait_download(self.path, f"{name}{datetime.now().strftime('%Y%m%d%H')}", data_download)
        time.sleep(1)
        file_path = os.path.join(self.path, file_name)
        data = pd.read_excel(file_path, header=1)
        rd = []
        for _, row in data.iterrows():
            rd.append(deal_func(row))
        os.remove(file_path)
        return rd


def wait_download(download_path, name, download_func):
    """等待开始下载"""
    files = os.listdir(download_path)
    for fname in files:
        if re.match(f"^{name}" + r".*\.xlsx$", fname) is None:
            continue
        print(f"清理下载前就存在的文件数据:{fname}")
        os.remove(os.path.join(download_path, fname))
    print(f"开始下载相关的文件信息:{name}")
    download_func()
    st = time.time()
    while True:
        if (time.time() - st) > 600:
            raise Exception("Waiting download timeout.")
        files = os.listdir(download_path)
        for fname in files:
            if re.match(f"^{name}" + r".*\.xlsx$", fname) is None:
                continue
            print(f"文件下载已完成:{fname}")
            return fname    


class _WebUrl:
    """网站地址"""

    def __init__(self) -> None:
        self.spfj = r"https://www.sinopharm-fj.com/spfj/flows/"  # 国控系网站
        self.inca = r"http://59.59.56.90:8094/ns/"  # 片仔癀漳州
        self.ly = r"http://www.luyan.com.cn/index.php"  # 鹭燕
        self.xm_tc = r"http://tc.tcyy.com.cn:8888/exm/login.jsp"  # 厦门同春
        self.fj_tc = r"http://tc.tcyy.com.cn:8888/etcyy/"  # 福建同春
        self.sm_tc = r"http://tc.tcyy.com.cn:8888/esmtc/"  # 三明同春
        self.druggc = r"https://zlcx.hrpzh.com/drugqc/home/login?type=flowLogin"  # 厦门片仔癀

WEBURL = _WebUrl()
