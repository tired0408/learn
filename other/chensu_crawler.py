"""
爬虫抓取医药相关的库存,并导出为xls表格
XLS属于较老版本,需使用xlwt数据库
"""
import re
import xlwt
import datetime
import collections
import pandas as pd
import numpy as np
import traceback
from xlwt.Worksheet import Worksheet
from selenium.common.exceptions import NoSuchElementException
from selenium.webdriver import Chrome
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.remote.webelement import WebElement
from selenium.webdriver.support.ui import WebDriverWait, Select
from selenium.webdriver.support import expected_conditions as EC


def correct_str(value):
    """修正字符串: 清理无用字符"""
    value = str(value)
    return value.strip()


class DataToExcel:
    """将数据转化为EXCEL表格类"""

    def __init__(self):
        self.database = self.get_production_database()
        self.data = collections.defaultdict(set)
        self.widths = [10, 10, 10, 12, 14]
        self.client_name = None
        self.date = datetime.datetime.today()
        self.date_style = xlwt.XFStyle()
        self.date_style.num_format_str = 'YYYY/MM/DD'

        self.row_i = 1
        self.wb = xlwt.Workbook()
        self.ws: Worksheet = self.init_sheet()

    def init_sheet(self) -> Worksheet:
        """初始化单元表"""
        rd: Worksheet = self.wb.add_sheet("商业库存导入模板")
        for j, value in enumerate(["一级商业*", "商品信息*", "本期库存*", "库存日期*", "库存获取日期*", "在途", "参考信息", "备注"]):
            rd.write(0, j, value)
        return rd

    def get_production_database(self):
        """获取产品数据库"""
        rd = collections.defaultdict(dict)
        database = pd.read_excel(r"E:\NewFolder\chensu\脚本产品库.xlsx")
        for _, row in database.iterrows():
            v0, v1, v2, v3, _ = row
            v3 = "" if pd.isna(v3) else v3
            rd[v0][v1] = [v2, v3]
        return rd

    def save(self):
        self.wb.save(r"E:\NewFolder\chensu\库存导入.xls")

    def append(self, name, count):
        """写入数据"""
        self.data[name].add(count)

    def len_byte(self, value):
        """获取字符串长度,一个中文的长度为2"""
        value = str(value)
        length = len(value)
        utf8_length = len(value.encode('utf-8'))
        length = (utf8_length - length) / 2 + length
        return int(length) + 2

    def write_to_excel(self):
        """将数据写入EXCEL表格"""
        self.widths[0] = max(self.widths[0], self.len_byte(self.client_name))
        name2standard = self.database[self.client_name]
        for product_name, number in self.data.items():
            if product_name in name2standard:
                product_name, reference = name2standard[product_name]
                remark = ""
            else:
                reference = ""
                remark = "未在产品库找到"
            number = int(np.sum(list(number)))
            self.ws.write(self.row_i, 0, self.client_name)
            self.ws.write(self.row_i, 1, product_name)
            self.ws.write(self.row_i, 2, number)
            self.ws.write(self.row_i, 3, self.date, self.date_style)
            self.ws.write(self.row_i, 4, self.date, self.date_style)
            self.ws.write(self.row_i, 6, reference)
            self.ws.write(self.row_i, 7, remark)
            self.row_i += 1

            self.widths[1] = max(self.widths[1], self.len_byte(product_name))
            self.widths[2] = max(self.widths[2], self.len_byte(str(number)))
        self.data.clear()
        print(f"[{self.client_name}]已将数据写入到excel表格中.")

    def cell_format(self):
        """设置单元格格式"""
        for i, width in enumerate(self.widths):
            self.ws.col(i).width = width * 256
        print("格式优化完成")


class Crawler:

    def __init__(self) -> None:
        self.driver = self.init_chrome()
        self.writer = DataToExcel()
        self.wait = WebDriverWait(self.driver, 5)
        self.spfj_url = r"https://www.sinopharm-fj.com/spfj/flows/"
        self.inca_url = r"http://59.59.56.90:8094/ns/"
        self.luyan_url = r"http://www.luyan.com.cn/index.php"
        self.tc_url = r"http://tc.tcyy.com.cn:8888/exm/login.jsp"
        self.druggc_url = r"http://117.29.176.58:8860/drugqc/home/login?type=flowLogin"

    def __del__(self):
        self.driver.quit()

    def init_chrome(self):
        exe_path = r'E:\py-workspace\learn\other\chromedriver.exe'
        service = Service(exe_path)
        options = Options()
        # options.add_argument("--headless")
        # options.add_argument("--disable-gpu")
        options.add_argument('--log-level=3')
        driver = Chrome(service=service, options=options)
        return driver

    def run(self, path):
        """运行脚本"""
        try:
            websites = pd.read_excel(path)
            # websites = websites[websites['流向查询网址'] != self.spfj_url]
            # websites = websites[websites['流向查询网址'] != self.inca_url]
            # websites = websites[websites['流向查询网址'] != self.luyan_url]
            self.writer.client_name = websites.iloc[0, 0]
            for _, row in websites.iterrows():
                client_name = correct_str(row.iloc[0])
                district_name = correct_str(row.iloc[1])
                website_url = correct_str(row.iloc[2])
                user = correct_str(row.iloc[3])
                password = row.iloc[4]
                password = "" if pd.isna(password) else correct_str(password)
                if client_name != self.writer.client_name:
                    self.writer.write_to_excel()
                    self.writer.client_name = client_name
                if website_url == self.spfj_url:
                    self.spfj_grab(user, password, district_name)
                elif website_url == self.inca_url:
                    self.inca_grab(user, password)
                elif website_url == self.luyan_url:
                    self.luyan_grab(user, password, district_name)
                elif website_url == self.tc_url:
                    self.tc_grab(user, password)
                elif website_url == self.druggc_url:
                    self.druggc_grab(user, password, district_name)
                else:
                    raise Exception("未定义该网站的爬虫抓取方法")
                break
            print("已完成所有数据写入,开始优化格式")
            self.writer.write_to_excel()
            self.writer.cell_format()
            self.writer.save()
            print("脚本已运行完成.")
        except Exception as e:
            print(traceback.format_exc())
            print(f"脚本运行出现异常: {e}")

    def spfj_grab(self, user, password, option_name):
        """国控福建流向查询系统的抓取
        Args:
            user: (str); 用户名
            password: (str); 密码
            option_name: (str); 选项名称
        """
        self.driver.get(self.spfj_url)
        self.wait.until(EC.visibility_of_element_located((By.ID, "login")))
        self.clear_and_send(self.driver.find_element(By.ID, "user"), user)
        self.clear_and_send(self.driver.find_element(By.ID, "pwd"), password)
        Select(self.driver.find_element(By.ID, "own")).select_by_visible_text(option_name)
        self.driver.find_element(By.ID, "login").click()
        self.wait.until(EC.visibility_of_element_located((By.ID, "butn")))
        self.driver.find_element(By.ID, "butn").click()
        WebDriverWait(self.driver, 30).until_not(EC.visibility_of_element_located((By.ID, "loading")))
        try:
            values = self.driver.find_element(By.ID, "customers")
            values = values.find_elements(By.TAG_NAME, "tr")
            for each_v in values[1:]:
                each_v = each_v.find_elements(By.TAG_NAME, "td")
                product_name = each_v[1].text + each_v[2].text
                inventory = int(each_v[8].text)
                self.writer.append(product_name, inventory)
            print(f"[{self.writer.client_name}]{user}的数据抓取已完成，共抓取{len(values) - 1}条数据")
        except NoSuchElementException:
            print(f"[{self.writer.client_name}]{user}的数据抓取已完成,共抓取0条数据")

    def inca_grab(self, user, password):
        """片仔癀漳州医药有限公司的数据抓取
        Args:
            user: (str); 用户名
            password: (str); 密码
        """
        module_name = '库存明细查询'
        self.driver.get(self.inca_url)
        self.wait.until(EC.visibility_of_element_located((By.ID, "login_link")))
        self.clear_and_send(self.driver.find_element(By.ID, "userName"), user)
        if password == "":
            self.driver.find_element(By.ID, "passWord").clear()
        else:
            self.clear_and_send(self.driver.find_element(By.ID, "passWord"), password)
        self.clear_and_send(self.driver.find_element(By.ID, "inputCode"),
                            self.driver.find_element(By.ID, "checkCode").text)
        self.wait.until(EC.visibility_of_element_located((By.TAG_NAME, "option")))
        self.driver.find_element(By.ID, "login_link").click()
        self.wait.until(EC.visibility_of_element_located((By.ID, "tree1")))
        tree = self.driver.find_element(By.ID, "tree1").find_element(By.TAG_NAME, "li")
        tree.find_element(By.CLASS_NAME, "l-expandable-close").click()
        tree.find_element(By.XPATH, f"//span[text()='{module_name}']").click()
        WebDriverWait(self.driver, 30).until(EC.visibility_of_element_located(
            (By.XPATH, f"//a[text()='{module_name}']/parent::*")))
        # 跳转到响应的iframe
        iframe_id = self.driver.find_element(By.XPATH, f"//a[text()='{module_name}']/parent::*")
        iframe_id = iframe_id.get_attribute("tabid")
        self.driver.switch_to.frame(self.driver.find_element(By.ID, iframe_id))
        try:
            # 获取表格
            total_num = 0
            while True:
                content = self.driver.find_element(By.CLASS_NAME, "bill_m")
                tabel = content.find_elements(By.CLASS_NAME, 'formsT_No_table')[1]
                values = tabel.find_elements(By.TAG_NAME, "tr")
                for each_v in values[1:]:
                    each_v = each_v.find_elements(By.TAG_NAME, "td")
                    product_name = each_v[1].text + each_v[2].text
                    inventory = int(each_v[5].text)
                    if "胰岛素" in product_name:
                        inventory = inventory * 2
                    total_num += 1
                    self.writer.append(product_name, inventory)
                pages = content.find_element(By.CLASS_NAME, "pages")
                page_info = pages.find_element(By.CLASS_NAME, "page_m").text
                index, total_index = page_info.split("/")
                if index == total_index:
                    break
                next_page = pages.find_element(By.CLASS_NAME, "next")
                next_page.click()
            print(f"[{self.writer.client_name}]{user}数据抓取已完成，共抓取{total_num}条数据")
        except NoSuchElementException:
            print(f"[{self.writer.client_name}]{user}数据抓取已完成,共抓取0条数据")

    def luyan_grab(self, user, password, district_name):
        """鹭燕医药网站的数据抓取
        Args:
            user: (str); 用户名
            password: (str); 密码
            district_name: (str); 区域名称
        """
        self.driver.get(self.luyan_url)
        self.wait.until(EC.visibility_of_element_located((By.CLASS_NAME, "buttonsubmit")))
        self.clear_and_send(self.driver.find_element(By.NAME, "username"), user)
        self.clear_and_send(self.driver.find_element(By.NAME, "loginpwd"), password)
        Select(self.driver.find_element(By.NAME, "select")).select_by_visible_text(district_name)
        self.driver.find_element(By.CLASS_NAME, "buttonsubmit").click()
        self.driver.switch_to.frame(self.driver.find_element(By.NAME, "menu"))
        self.driver.find_element(By.XPATH, "//strong[text()='库存明细信息']").click()
        self.driver.switch_to.default_content()
        self.driver.switch_to.frame(self.driver.find_element(By.NAME, "main"))
        try:
            values = self.driver.find_element(By.CLASS_NAME, "gridBody").find_elements(By.TAG_NAME, "tr")
            for each_v in values[1:]:
                each_v = each_v.find_elements(By.TAG_NAME, "td")
                product_name = each_v[0].text + each_v[1].text
                inventory = int(float(each_v[3].text))
                self.writer.append(product_name, inventory)
            print(f"[{self.writer.client_name}]数据抓取已完成，共抓取{len(values) - 1}条数据")
        except NoSuchElementException:
            print(f"[{self.writer.client_name}]{user}数据抓取已完成,共抓取0条数据")

    def tc_grab(self, user, password):
        """同春医药网站的数据抓取
        Args:
            user: (str); 用户名
            password: (str); 密码
        """
        self.driver.get(self.tc_url)
        self.wait.until(EC.visibility_of_element_located((By.ID, "txt_UserName")))
        self.clear_and_send(self.driver.find_element(By.ID, "txt_UserName"), user)
        self.clear_and_send(self.driver.find_element(By.ID, "txtPassWord"), password)
        WebDriverWait(self.driver, 120).until(EC.visibility_of_element_located((By.XPATH, "//table[@border='1']")))
        try:
            values = self.driver.find_element(By.XPATH, "//table[@border='1']").find_elements(By.TAG_NAME, "tr")
            for each_v in values[1:]:
                each_v = each_v.find_elements(By.TAG_NAME, "td")
                product_name = each_v[1].text + each_v[2].text
                inventory = int(each_v[6].text)
                self.writer.append(product_name, inventory)
            print(f"[{self.writer.client_name}]数据抓取已完成，共抓取{len(values) - 1}条数据")
        except NoSuchElementException:
            print(f"[{self.writer.client_name}]数据抓取已完成,共抓取0条数据")

    def druggc_grab(self, user, password, district_name):
        """厦门片仔癀宏仁医药有限公司网站的数据抓取
        Args:
            user: (str); 用户名
            password: (str); 密码
            district_name: (str); 区域名称
        """
        self.driver.get(self.druggc_url)
        self.wait.until(EC.visibility_of_element_located((By.ID, "username")))
        self.clear_and_send(self.driver.find_element(By.ID, "username"), user)
        self.clear_and_send(self.driver.find_element(By.ID, "password"), password)
        Select(self.driver.find_element(By.ID, "entryid")).select_by_visible_text(district_name)
        WebDriverWait(self.driver, 120).until(EC.visibility_of_element_located((By.XPATH, "//span[text()='库存明细']")))
        self.driver.find_element(By.XPATH, "//span[text()='库存明细']").click()
        self.driver.switch_to.frame(self.driver.find_element(By.ID, "mainframe"))
        try:
            total_num = 0
            while True:
                self.wait.until_not(EC.visibility_of_element_located((By.CLASS_NAME, "fixed-table-loading")))
                values = self.driver.find_element(By.ID, "dataTable")
                values = values.find_elements(By.TAG_NAME, "tr")
                for each_v in values[1:]:
                    each_v = each_v.find_elements(By.TAG_NAME, "td")
                    product_name = each_v[0].text + each_v[2].text
                    inventory = int(each_v[-2].text)
                    total_num += 1
                    self.writer.append(product_name, inventory)
                page_info = self.driver.find_element(By.CLASS_NAME, "fixed-table-pagination")
                page_numbers = re.findall(r'\d+', page_info.find_element(By.CLASS_NAME, "pagination-info").text)
                if page_numbers[-1] == page_numbers[-2]:
                    break
                page_info.find_element(By.XPATH, "//li[contains(@class, 'page-next')]/a").click()
            print(f"[{self.writer.client_name}]数据抓取已完成，共抓取{total_num}条数据")
        except NoSuchElementException:
            print(f"[{self.writer.client_name}]{user}数据抓取已完成,共抓取0条数据")

    def clear_and_send(self, element: WebElement, value):
        """清除input框内的值并输入所需数据"""
        element.clear()
        element.send_keys(value)


if __name__ == "__main__":
    crawler = Crawler()
    crawler.run(r"E:\NewFolder\chensu\库存网查明细.xlsx")
