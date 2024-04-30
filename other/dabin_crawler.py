"""
爬虫抓取大饼所需的进销存报表,并导出为xlsx表格
接受OCR识别请求地址,  http://localhost:8557/ocr
"""
import re
import openpyxl
import traceback
import collections
import pandas as pd
import numpy as np
from re import Match
from tqdm import tqdm
from datetime import datetime, timedelta
from openpyxl.cell.cell import Cell
from openpyxl.styles import PatternFill
from selenium.webdriver.support.ui import WebDriverWait
from .medicine_utils import start_http, init_chrome, SPFJWeb, DruggcWeb, LYWeb, INCAWeb, correct_str


class Data:

    def __init__(self) -> None:
        self.purchase = collections.defaultdict(set)  # 本周进货
        self.inventory = collections.defaultdict(set)  # 本周库存
        self.code = collections.defaultdict(set)  # 批号
        self.sales = collections.defaultdict(set)  # 本周销量


class DataToExcel:
    """读取xlsx文件,并修改里面的数值"""

    def __init__(self, path, standard_path, interval=1) -> None:
        self.path = path
        self.wb = openpyxl.load_workbook(self.path)
        self.ws = self.create_ws(interval)
        self.data = collections.defaultdict(Data)
        self.name2standard = self.init_standard_data(standard_path)

    def init_standard_data(self, path):
        return {
            "国药控股福建有限公司": {},
            "鹭燕医药股份有限公司": {},
            "泉州鹭燕医药有限公司": {},
        }

    def save(self):
        """保存文件"""
        self.wb.save(self.path)

    def create_ws(self, interval):
        """创建工作表"""
        print("复制工作表")
        last_week = self.wb.worksheets[0]
        ws = self.wb.copy_worksheet(last_week)
        week_pattern: Match = re.search(r"^(\d+)月(\d+)周$", last_week.title)
        month, week = week_pattern.groups()
        week = int(week) + interval
        ws.title = f"{int(month) + 1}月1周" if week > 4 else f"{month}月{week}周"
        self.wb.move_sheet(ws, offset=-self.wb.index(ws))
        print("删除颜色")
        for row in ws.iter_rows():
            for cell in row:
                assert isinstance(cell, Cell)
                cell.fill = PatternFill(start_color="FFFFFF", end_color="FFFFFF", fill_type="solid")
        self.save()
        return ws

    def write_to_excel(self):
        """将数据写入excel表格"""
        now_date = datetime.now().strftime("%Y/%m/%d")
        for row_index in tqdm(range(3, self.ws.max_row)):
            name = self.ws[f"G{row_index}"].value
            person = self.ws[f"E{row_index}"].value
            company = self.ws[f"F{row_index}"].value
            data = self.data[company]
            self.ws[f"A{row_index}"] = now_date  # 日期
            if person == "傅镔滢":
                self.ws[f"L{row_index}"] = self.ws[f"Q{row_index}"].value  # 上周库存
                self.ws[f"M{row_index}"] = int(np.sum(list(data.purchase[name])))  # 本周进货
                self.ws[f"Q{row_index}"] = int(np.sum(list(data.inventory[name])))  # 本周库存
                self.ws[f"R{row_index}"] = "、".join(list(data.code[name])) if name in data.code else ""  # 批号
                self.ws[f"S{row_index}"] = self.judge_worn(row_index, int(np.sum(list(data.sales[name]))))  # 备注（记录破损情况）
            else:
                self.ws[f"L{row_index}"] = ""
                self.ws[f"M{row_index}"] = ""
                self.ws[f"Q{row_index}"] = ""
                self.ws[f"R{row_index}"] = ""
                self.ws[f"S{row_index}"] = ""
        self.save()

    def judge_worn(self, row_index, actual_sale):
        """判断是否有破损"""
        theory_sale = self.ws[f"L{row_index}"].value + self.ws[f"M{row_index}"].value + \
            self.ws[f"N{row_index}"].value - self.ws[f"P{row_index}"].value - self.ws[f"Q{row_index}"].value
        if theory_sale != actual_sale:
            return theory_sale - actual_sale
        else:
            return ""


def main(websites_path, chrome_exe_path, save_path, database_path, week_interval):
    http_server, q = start_http()
    driver = init_chrome(chrome_exe_path)
    wait = WebDriverWait(driver, 5)
    writer = DataToExcel(save_path, database_path, interval=week_interval)
    spfj = SPFJWeb(driver, wait)
    druggc = DruggcWeb(driver, wait, q)
    luyan = LYWeb(driver, wait)
    inca = INCAWeb(driver, wait)

    # 计算开始时间
    start_date = datetime.now()
    start_date = datetime.strptime("2024-04-14", "%Y-%m-%d")
    start_date = start_date - timedelta(days=start_date.weekday() + 7 * (week_interval - 1))
    start_date_str = start_date.strftime("%Y-%m-%d")
    websites = pd.read_excel(websites_path)
    for _, row in websites.iterrows():
        # 获取登录信息
        client = correct_str(row.iloc[3])
        district_name = correct_str(row.iloc[4])
        website_url = correct_str(row.iloc[7])
        user = correct_str(row.iloc[8])
        password = row.iloc[9]
        password = "" if pd.isna(password) else correct_str(password)
        # 获取相关数据
        client_data = writer.data[client]
        name2standard = writer.name2standard[client]
        if website_url == spfj.url:
            spfj.login(user, password, district_name)
            for product_name, purchase, sale, inventory in spfj.purchase_sale_stock(start_date_str):
                standard = name2standard[product_name]
                client_data.purchase[standard].add(purchase)
                client_data.sales[standard].add(sale)
                client_data.inventory[standard].add(inventory)
            for product_name, _, code in spfj.get_inventory(start_date_str):
                standard = name2standard[product_name]
                client_data.code[standard].add(code)
        elif website_url == luyan.url:
            luyan.login(user, password, district_name)
            for product_name, _, code in luyan.get_inventory():
                standard = name2standard[product_name]
                client_data.code[standard].add(code)
            for product_name, purchase, sale, inventory in luyan.purchase_sale_stock(start_date_str):
                standard = name2standard[product_name]
                client_data.purchase[standard].add(purchase)
                client_data.sales[standard].add(sale)
                client_data.inventory[standard].add(inventory)
        elif website_url == inca.url:
            inca.login(user, password)
            for product_name, inventory, code in inca.get_inventory():
                standard = name2standard[product_name]
                client_data.inventory[standard].add(inventory)
                client_data.code[standard].add(code)
            for product_name, amount in inca.get_purchase(start_date_str):
                standard = name2standard[product_name]
                client_data.purchase[standard].add(amount)
            for product_name, amount in inca.get_sales(start_date_str):
                standard = name2standard[product_name]
                client_data.sales[standard].add(amount)
        elif website_url == druggc.url:
            druggc.login(user, password, district_name)
            for product_name, inventory, code in druggc.get_inventory():
                standard = name2standard[product_name]
                client_data.inventory[standard].add(inventory)
                client_data.code[standard].add(code)
            for product_name, amount in druggc.get_purchase(start_date_str):
                standard = name2standard[product_name]
                client_data.purchase[standard].add(amount)
            for product_name, amount in druggc.get_sales(start_date_str):
                standard = name2standard[product_name]
                client_data.sales[standard].add(amount)
        else:
            raise Exception("未知网站")
    print("将数据写入到excel表格中")
    writer.write_to_excel()
    print("关闭HTTP服务器")
    http_server.close_server()


if __name__ == "__main__":
    try:
        set_websites_path = r"E:\NewFolder\dabin\福建商业明细表(福建)22.2.10-主席.xlsx"
        set_chrome_exe_path = r'E:\NewFolder\chromedriver_mac_arm64_114\chromedriver.exe'
        set_save_path = r"E:\NewFolder\dabin\data.xlsx"
        set_database_path = r"E:\NewFolder\dabin\standard.xlsx"
        set_week_interval = 1  # 间隔的查询周数
        main(set_websites_path, set_chrome_exe_path, set_save_path, set_database_path, set_week_interval)
        print("脚本已运行完成.")
    except Exception:
        print("-" * 150)
        print("脚本运行出现异常:")
        print(traceback.format_exc())
        print("-" * 150)
