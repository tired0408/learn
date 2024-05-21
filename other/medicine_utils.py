"""
医药网站抓取的通用方法
"""
import re
import time
import socket
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


def init_chrome(chromedriver_path, chrome_path=None, is_proxy=True):
    """初始化浏览器"""
    service = Service(chromedriver_path)
    options = Options()
    if chrome_path is not None:
        options.binary_location = chrome_path
    if is_proxy:
        options.add_argument('--proxy-server=127.0.0.1:8080')
        options.add_argument('ignore-certificate-errors')
    options.add_argument('--log-level=3')
    driver = Chrome(service=service, options=options)
    return driver


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
    """分析网站数据"""
    websites = pd.read_excel(path)
    websites_by_code, websites_no_code = [], []
    # 网站的验证码情况
    code_condition = {
        SPFJWeb.url: False,
        INCAWeb.url: False,
        LYWeb.url: False,
        TCWeb.url: True,
        DruggcWeb.url: True
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
        if ignore_names is not None and data["client_name"] in ignore_names:
            print(f"断点之前已查询过:{data['client_name']},{data['user']}")
            continue
        if code_condition[data["website_url"]]:
            websites_by_code.append(data)
        else:
            websites_no_code.append(data)
    return websites_by_code, websites_no_code


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


class SPFJWeb:
    """国控福建网站的数据抓取"""
    url = r"https://www.sinopharm-fj.com/spfj/flows/"

    def __init__(self, driver) -> None:
        self.driver: Chrome = driver

    def login(self, user, password, district_name):
        get_url_success(self.driver, self.url, By.ID, "login")
        clear_and_send(self.driver.find_element(By.ID, "user"), user)
        clear_and_send(self.driver.find_element(By.ID, "pwd"), password)
        Select(self.driver.find_element(By.ID, "own")).select_by_visible_text(district_name)
        self.driver.find_element(By.ID, "login").click()
        WebDriverWait(self.driver, 10).until(EC.visibility_of_element_located((By.ID, "butn")))
        print(f"[国控福建]{user}用户已登录")

    def purchase_sale_stock(self, start_date=None):
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
        rd = self.get_table_data(deal_func, "进销存汇总", start_date)
        print(f"[国控福建]进销存数据抓取已完成，共抓取{len(rd)}条数据")
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
        print(f"[国控福建]库存数据抓取已完成，共抓取{len(rd)}条数据")
        return rd

    def get_table_data(self, deal_func, table_type, start_date=None):
        """获取表格数据的通用方法"""
        Select(self.driver.find_element(By.ID, "type")).select_by_visible_text(table_type)
        if start_date is not None:
            start_element = self.driver.find_element(By.ID, "txtBeginDate")
            self.driver.execute_script("arguments[0].value = arguments[1]", start_element, start_date)
        self.driver.find_element(By.ID, "butn").click()
        WebDriverWait(self.driver, 30).until_not(EC.visibility_of_element_located((By.ID, "loading")))
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
    url = r"http://59.59.56.90:8094/ns/"

    def __init__(self, driver) -> None:
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
        WebDriverWait(self.driver, 10).until(EC.visibility_of_element_located((By.ID, "tree1")))
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
            if "胰岛素" in product_name:
                amount = amount * 2
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
        rd = self.get_table_data(deal_func, "供应商网络服务", "进货明细查询", start_date, self.time_set)
        print(f"[片仔癀漳州]进货数据抓取已完成，共抓取{len(rd)}条数据")
        return rd

    def get_sales(self, start_date=None):
        def deal_func(elements: List[WebElement]):
            product_name = elements[2].text + elements[3].text
            product_name = product_name.replace(" ", "")
            amount = int(elements[9].text)
            return [product_name, amount]
        rd = self.get_table_data(deal_func, "客户网络服务", "发货明细查询", start_date, self.time_set)
        print(f"[片仔癀漳州]销售数据抓取已完成，共抓取{len(rd)}条数据")
        return rd

    def time_set(self, start_date):
        """时间设置"""
        self.driver.find_element(By.ID, "but_b").click()
        WebDriverWait(self.driver, 10).until(EC.visibility_of_element_located((By.ID, "but_con")))
        self.driver.find_element(By.ID, "but_con").click()
        WebDriverWait(self.driver, 10).until(EC.visibility_of_element_located((By.CLASS_NAME, "modal-content")))
        self.driver.find_element(By.XPATH, "//input[@value='shijian']").click()
        start_element = self.driver.find_element(By.NAME, "startdate")
        self.driver.execute_script("arguments[0].value = arguments[1]", start_element, start_date)
        end_element = self.driver.find_element(By.NAME, "enddate")
        end_date = datetime.now().strftime("%Y-%m-%d")
        self.driver.execute_script("arguments[0].value = arguments[1]", end_element, end_date)
        self.driver.find_element(By.XPATH, "//div[contains(@class, 'modal-footer')]/input[@value='查询']").click()

    def get_table_data(self, deal_func, tree_type, table_type, start_date=None, time_func=None):
        """获取表格数据"""
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
        if start_date is None:
            self.driver.find_element(By.ID, "submit_s").click()
        else:
            time_func(start_date)
        # 获取表格
        WebDriverWait(self.driver, 30).until_not(EC.visibility_of_element_located((By.CLASS_NAME, "l-tab-loading")))
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


class LYWeb:
    """鹭燕网站的数据抓取"""
    url = r"http://www.luyan.com.cn/index.php"

    def __init__(self, driver) -> None:
        self.driver: Chrome = driver

    def login(self, user, password, district_name):
        get_url_success(self.driver, self.url, By.CLASS_NAME, "buttonsubmit")
        clear_and_send(self.driver.find_element(By.NAME, "username"), user)
        clear_and_send(self.driver.find_element(By.NAME, "loginpwd"), password)
        Select(self.driver.find_element(By.NAME, "select")).select_by_visible_text(district_name)
        self.driver.find_element(By.CLASS_NAME, "buttonsubmit").click()
        WebDriverWait(self.driver, 10).until(EC.visibility_of_element_located((By.NAME, "menu")))
        print(f"[鹭燕]{user}用户已登录")

    def get_inventory(self):
        """
        获取库存数据
        :return: [(商品名称, 库存数量, 批号), ...]
        """
        def deal_func(elements: List[WebElement]):
            product_name = elements[0].text + elements[1].text
            product_name = product_name.replace(" ", "")
            amount = int(float(elements[3].text))
            code = str(elements[5].text)
            return [product_name, amount, code]
        rd = self.get_table_data(deal_func, "库存明细信息")
        print(f"[鹭燕]库存数据抓取已完成，共抓取{len(rd)}条数据")
        return rd

    def purchase_sale_stock(self, start_date):
        """
        进销存数据抓取
        :return: [(商品名称, 进货数量, 销售数量, 库存数量), ...]
        """
        def deal_func(elements: List[WebElement]):
            product_name = elements[0].text + elements[1].text
            product_name = product_name.replace(" ", "")
            purchase = int(float(elements[4].text))
            sales = int(float(elements[3].text))
            inventory = int(float(elements[5].text))
            return [product_name, purchase, sales, inventory]
        rd = self.get_table_data(deal_func, "进销存汇总表", start_date)
        print(f"[鹭燕]进销存数据抓取已完成，共抓取{len(rd)}条数据")
        return rd

    def get_table_data(self, deal_func, table_type, start_date=None):
        """获取表格数据"""
        self.driver.switch_to.default_content()
        self.driver.switch_to.frame(self.driver.find_element(By.NAME, "menu"))
        self.driver.find_element(By.XPATH, f"//strong[text()='{table_type}']").click()
        self.driver.switch_to.default_content()
        self.driver.switch_to.frame(self.driver.find_element(By.NAME, "main"))
        if start_date is not None:
            start_element = self.driver.find_element(By.NAME, "StartDate")
            self.driver.execute_script("arguments[0].value = arguments[1]", start_element, start_date)
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
    url = r"http://tc.tcyy.com.cn:8888/exm/login.jsp"

    def __init__(self, driver, captcha) -> None:
        self.driver: Chrome = driver
        self.captcha: Queue = captcha

    def login(self, user, password):
        get_url_success(self.driver, self.url, By.ID, "imgLogin")
        while True:
            try:
                WebDriverWait(self.driver, 10).until(EC.visibility_of_element_located((By.ID, "imgLogin")))
                clear_and_send(self.driver.find_element(By.ID, "txt_UserName"), user)
                clear_and_send(self.driver.find_element(By.ID, "txtPassWord"), password)
                captcha_value = self.captcha.get(timeout=60)
                print(f"输入验证码:{captcha_value}")
                clear_and_send(self.driver.find_element(By.ID, "txtVerifyCode"), captcha_value)
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
            pass
        print(f"[同春医药]库存数据抓取已完成，共抓取{len(rd)}条数据")
        return rd


class DruggcWeb:
    """片仔癀宏仁医药有限公司网站的数据抓取"""
    url = r"http://117.29.176.58:8860/drugqc/home/login?type=flowLogin"

    def __init__(self, driver, captcha) -> None:
        self.driver: Chrome = driver
        self.captcha: Queue = captcha

    def login(self, user, password, district_name):
        get_url_success(self.driver, self.url, By.ID, "login")
        clear_and_send(self.driver.find_element(By.ID, "username"), user)
        clear_and_send(self.driver.find_element(By.ID, "password"), password)
        Select(self.driver.find_element(By.ID, "entryid")).select_by_visible_text(district_name)
        while True:
            captcha = self.captcha.get(timeout=60)
            print(f"输入验证码:{captcha}")
            if len(captcha) == 4:
                clear_and_send(self.driver.find_element(By.ID, "captcha"), captcha)
                self.driver.find_element(By.ID, "login").click()
                captcha_tip = EC.visibility_of_element_located((By.XPATH, "//div[text()='验证码不正确！']"))
                try:
                    WebDriverWait(self.driver, 10).until(captcha_tip)
                except TimeoutException:
                    break
                WebDriverWait(self.driver, 5).until_not(captcha_tip)
            else:
                print("[片仔癀宏仁医药]验证码识别错误，更换验证码图片")
            self.driver.find_element(By.ID, "captchaImg").click()
        WebDriverWait(self.driver, 10).until(EC.visibility_of_element_located((By.ID, "side-menu")))
        print(f"[片仔癀宏仁医药]{user}用户已登录")

    def get_inventory(self):
        """
        获取库存数据
        :return: [(商品名称, 库存数量, 批号), ...]
        """
        def deal_inventory(elements: List[WebElement]):
            product_name = elements[0].text + elements[2].text
            product_name = product_name.replace(" ", "")
            amount = int(elements[-2].text)
            code = str(elements[4].text)
            return [product_name, amount, code]
        rd = self.get_table_data(deal_inventory, "库存明细")
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
        rd = self.get_table_data(deal_restock, "进货明细", start_date)
        print(f"[片仔癀宏仁医药]进货数据抓取已完成，共抓取{len(rd)}条数据")
        return rd

    def get_sales(self, start_date):
        """
        获取销售数据
        :return: [(商品名称, 销售数量), ...]
        """
        def deal_sales(elements: List[WebElement]):
            product_name = elements[2].text + elements[3].text
            product_name = product_name.replace(" ", "")
            amount = int(elements[4].text)
            return [product_name, amount]
        rd = self.get_table_data(deal_sales, '供应商流向', start_date)
        print(f"[片仔癀宏仁医药]销售数据抓取已完成，共抓取{len(rd)}条数据")
        return rd

    def get_table_data(self, deal_func, table_type, start_date=None):
        """获取表格数据"""
        self.driver.switch_to.default_content()
        self.driver.find_element(By.XPATH, f"//span[text()='{table_type}']").click()
        self.driver.switch_to.frame(self.driver.find_element(By.ID, "mainframe"))
        if start_date is not None:
            start_element = self.driver.find_element(By.ID, "beginCreateTime")
            self.driver.execute_script("arguments[0].value = arguments[1]", start_element, start_date)
        try:
            self.driver.find_element(By.XPATH, "//button[text()='查询']").click()
        except NoSuchElementException:
            pass
        rd = []
        while True:
            condition = EC.visibility_of_element_located((By.CLASS_NAME, "fixed-table-loading"))
            WebDriverWait(self.driver, 10).until_not(condition)
            # 网站有个BUG，数据显示太慢，会先显示没有匹配到数据，然后再显示数据
            try:
                self.driver.find_element(By.CLASS_NAME, "no-records-found")
                condition = EC.visibility_of_element_located((By.CLASS_NAME, "no-records-found"))
                WebDriverWait(self.driver, 10).until_not(condition)
            except NoSuchElementException:
                pass
            except TimeoutException:
                break
            values = self.driver.find_element(By.ID, "dataTable")
            values = values.find_elements(By.TAG_NAME, "tr")
            for each_v in values[1:]:
                each_v = each_v.find_elements(By.TAG_NAME, "td")
                rd.append(deal_func(each_v))
            page_info = self.driver.find_element(By.CLASS_NAME, "fixed-table-pagination")
            page_numbers = re.findall(r'\d+', page_info.find_element(By.CLASS_NAME, "pagination-info").text)
            if page_numbers[-1] == page_numbers[-2]:
                break
            page_info.find_element(By.XPATH, "//li[contains(@class, 'page-next')]/a").click()
        return rd
