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
from typing import Dict, List, Tuple
from datetime import datetime, timedelta
from openpyxl.cell.cell import Cell
from openpyxl.styles import PatternFill
from openpyxl.worksheet.worksheet import Worksheet
from openpyxl.utils.cell import column_index_from_string
from medicine_utils import start_socket, init_chrome, analyze_website, SPFJWeb, DruggcWeb, LYWeb, INCAWeb


class ClientData:
    """每个客户的数据类"""

    def __init__(self) -> None:
        self.purchase = collections.defaultdict(list)  # 本周进货
        self.inventory = collections.defaultdict(list)  # 本周库存
        self.code = collections.defaultdict(list)  # 批号
        self.sales = collections.defaultdict(list)  # 本周销量


class DataToExcel:
    """读取xlsx文件,并修改里面的数值"""

    def __init__(self, path, interval=1) -> None:
        self.wb = openpyxl.load_workbook(path)
        self.ws, self.path = self.create_ws(interval, path)

    def __del__(self):
        self.wb.close()

    def save(self):
        """保存文件"""
        self.wb.save(self.path)

    def create_ws(self, interval, path: str) -> Tuple[Worksheet, str]:
        """创建工作表"""
        last_ws = self.wb.worksheets[0]
        now_date = datetime.now().strftime("%Y/%m/%d")
        if last_ws.cell(3, 1).value.strftime("%Y/%m/%d") == now_date:
            return last_ws
        print("复制工作表")
        ws = self.wb.copy_worksheet(last_ws)
        week_pattern: Match = re.search(r"^(\d+)月(\d+)周$", last_ws.title)
        last_month, last_week = week_pattern.groups()
        if int(last_week) > 4 - interval:
            month = int(last_month) + 1
            week = 1
        else:
            month = last_month
            week = int(last_week) + interval
        ws.title = f"{month}月{week}周"
        self.wb.move_sheet(ws, offset=-self.wb.index(ws))
        print("删除颜色")
        for row in ws.iter_rows():
            for cell in row:
                assert isinstance(cell, Cell)
                cell.fill = PatternFill(start_color="FFFFFF", end_color="FFFFFF", fill_type="solid")
        path = path.replace(f"2024年{last_month}月第{last_week}周", f"2024年{month}月第{week}周")
        return ws, path

    def write_to_excel(self, all_data: Dict[str, ClientData]):
        """将数据写入excel表格"""
        now_date = datetime.now()
        for row_i in tqdm(range(3, self.ws.max_row + 1)):
            name = self.ws.cell(row_i, column_index_from_string("G")).value
            person = self.ws.cell(row_i, column_index_from_string("E")).value
            company = self.ws.cell(row_i, column_index_from_string("F")).value
            # 写入数据
            statistics_date = self.ws.cell(row_i, column_index_from_string("A"))
            if statistics_date.value.strftime("%Y/%m/%d") != now_date.strftime("%Y/%m/%d"):
                statistics_date.value = now_date  # 统计日期
                last_inventory = self.ws.cell(row_i, column_index_from_string("Q")).value
                self.ws.cell(row_i, column_index_from_string("L"), last_inventory)  # 上周库存
            if person == "傅镔滢" and company in all_data:
                data = all_data[company]
                this_purchase = int(np.sum(list(data.purchase[name])))
                self.ws.cell(row_i, column_index_from_string("M"), this_purchase)
                this_inventory = int(np.sum(list(data.inventory[name])))
                self.ws.cell(row_i, column_index_from_string("Q"), this_inventory)
                this_code = "、".join(list(data.code[name])) if name in data.code else ""
                self.ws.cell(row_i, column_index_from_string("R"), this_code)
                this_worn = self.judge_worn(row_i, int(np.sum(list(data.sales[name]))))
                self.ws.cell(row_i, column_index_from_string("S"), this_worn)
            else:
                self.ws.cell(row_i, column_index_from_string("M"), "")  # 本周进货
                self.ws.cell(row_i, column_index_from_string("Q"), "")  # 本周库存
                self.ws.cell(row_i, column_index_from_string("R"), "")  # 批号
                self.ws.cell(row_i, column_index_from_string("S"), "")  # 备注（记录破损情况）

    def judge_worn(self, row_index, actual_sale):
        """判断是否有破损"""
        theory_sale = self.ws[f"L{row_index}"].value + self.ws[f"M{row_index}"].value + \
            self.ws[f"N{row_index}"].value - self.ws[f"P{row_index}"].value - self.ws[f"Q{row_index}"].value
        if theory_sale != actual_sale:
            return theory_sale - actual_sale
        else:
            return ""


class SPFJWebSelf(SPFJWeb):

    def get_datas(self, user, password, district_name, start_date_str, save_data: ClientData, name2standard):
        """获取所需数据"""
        self.login(user, password, district_name)
        for product_name, purchase, sale, inventory in self.purchase_sale_stock(start_date_str):
            standard = name2standard[product_name]
            save_data.purchase[standard].append(purchase)
            save_data.sales[standard].append(sale)
            save_data.inventory[standard].append(inventory)
        for product_name, _, code in self.get_inventory(start_date_str):
            standard = name2standard[product_name]
            save_data.code[standard].append(code)


class LYWebSelf(LYWeb):

    def get_datas(self, user, password, district_name, start_date_str, save_data: ClientData, name2standard):
        self.login(user, password, district_name)
        for product_name, _, code in self.get_inventory():
            standard = name2standard[product_name]
            save_data.code[standard].append(code)
        for product_name, purchase, sale, inventory in self.purchase_sale_stock(start_date_str):
            standard = name2standard[product_name]
            save_data.purchase[standard].append(purchase)
            save_data.sales[standard].append(sale)
            save_data.inventory[standard].append(inventory)


class DruggcWebSelf(DruggcWeb):

    def get_datas(self, user, password, district_name, start_date_str, save_data: ClientData, name2standard):
        self.login(user, password, district_name)
        for product_name, inventory, code in self.get_inventory():
            standard = name2standard[product_name]
            save_data.inventory[standard].append(inventory)
            save_data.code[standard].append(code)
        for product_name, amount in self.get_purchase(start_date_str):
            standard = name2standard[product_name]
            save_data.purchase[standard].append(amount)
        for product_name, amount in self.get_sales(start_date_str):
            standard = name2standard[product_name]
            save_data.sales[standard].append(amount)


class INCAWebSelf(INCAWeb):

    def get_datas(self, user, password, start_date_str, save_data: ClientData, name2standard):
        self.login(user, password)
        for product_name, inventory, code in self.get_inventory():
            standard = name2standard[product_name]
            save_data.inventory[standard].append(inventory)
            save_data.code[standard].append(code)
        for product_name, amount in self.get_purchase(start_date_str):
            standard = name2standard[product_name]
            save_data.purchase[standard].append(amount)
        for product_name, amount in self.get_sales(start_date_str):
            standard = name2standard[product_name]
            save_data.sales[standard].append(amount)


def gain_breakpoint(path):
    """读取断点数据，获取已经爬取的客户名称"""
    df = pd.read_excel(path, header=None)
    if df.iloc[3, 0].strftime("%Y/%m/%d") != datetime.now().strftime("%Y/%m/%d"):
        print("无断点数据，从头开始抓取")
        return
    df = df.iloc[2:]
    df = df[df[4] == "傅镔滢"]
    df = df.iloc[:, [5, 12, 16, 17, 18]]
    df = df.fillna("")
    df[12] = df[12].astype(str)
    df[16] = df[16].astype(str)
    df[17] = df[17].astype(str)
    df[18] = df[18].astype(str)
    df = df.groupby(5).agg({12: 'sum', 16: 'sum', 17: 'sum', 18: 'sum'})
    df = df.where(df != "", np.nan)
    df = df.dropna(subset=[12, 16, 17, 18], how="all")
    rd = df.index.to_list()
    print(f"检测到断点,进行断点续查。已抓取数据:{rd}")
    return rd


def gain_database(path):
    """获取数据库"""
    rd = collections.defaultdict(dict)
    database = pd.read_excel(path)
    for _, row in database.iterrows():
        client_production: str = row[1]
        if pd.isna(client_production):
            continue
        client_production = client_production.replace(" ", "")
        rd[row[0]][client_production] = row[2]
    return rd


def crawler_from_web(chrome_path, chromedriver_path, database_path, websites_by_code, websites_no_code,
                     week_interval) -> Dict[str, ClientData]:
    """从网站上爬取数据"""
    def crawler_general(datas: List[dict], url2method):
        for data in datas:
            client_name = data.pop("client_name")
            website_url = data.pop("website_url")
            try:
                data["start_date_str"] = start_date_str
                data["save_data"] = rd[client_name]
                data["name2standard"] = all_name2standard[client_name]
                url2method[website_url](**data)
            except Exception:
                print("-" * 150)
                print(f"脚本运行出现异常, 出错的截至问题公司:{client_name}")
                print(traceback.format_exc())
                print("-" * 150)
                print("去除该客户的全部数据")
                rd.pop(client_name)
                return True
        return False

    rd = collections.defaultdict(ClientData)
    print("获取数据库")
    all_name2standard = gain_database(database_path)
    print("计算开始时间")
    start_date = datetime.now()
    start_date = start_date - timedelta(days=start_date.weekday() + 7 * (week_interval - 1))
    start_date_str = start_date.strftime("%Y-%m-%d")
    if len(websites_no_code) != 0:
        print("抓取无验证码的网站数据")
        driver = init_chrome(chromedriver_path, chrome_path=chrome_path, is_proxy=False)
        spfj = SPFJWebSelf(driver)
        inca = INCAWebSelf(driver)
        luyan = LYWebSelf(driver)
        url_condition = {
            spfj.url: spfj.get_datas,
            inca.url: inca.get_datas,
            luyan.url: luyan.get_datas
        }
        is_error = crawler_general(websites_no_code, url_condition)
        print("关闭浏览器")
        driver.quit()
        if is_error:
            return rd
    if len(websites_by_code) != 0:
        print("抓取有验证码的网站数据")
        q = start_socket()
        driver = init_chrome(chromedriver_path, chrome_path=chrome_path)
        druggc = DruggcWebSelf(driver, q)
        url_condition = {
            druggc.url: druggc.get_datas
        }
        print("开始抓取")
        crawler_general(websites_by_code, url_condition)
        print("关闭浏览器")
        driver.quit()
    return rd


def main(chrome_path, chromedriver_path, websites_path, save_path, database_path, week_interval):
    print("读取断点数据")
    breakpoint_names = gain_breakpoint(save_path)
    print("针对网站数据进行分类")
    websites_by_code, websites_no_code = analyze_website(websites_path, breakpoint_names)
    print("从网站上爬取数据")
    crawler_data = crawler_from_web(chrome_path, chromedriver_path, database_path,
                                    websites_by_code, websites_no_code, week_interval)
    print("定义数据写入类")
    writer = DataToExcel(save_path, interval=week_interval)
    print("将数据写入到excel表格中")
    writer.write_to_excel(crawler_data)
    writer.save()
    print("程序运行已完成")


if __name__ == "__main__":
    set_chromedriver_path = r'E:\NewFolder\chromedriver_mac_arm64_114\chromedriver.exe'
    set_chrome_path = r"E:\NewFolder\chromedriver_mac_arm64_114\chrome114\App\Chrome-bin\chrome.exe"

    set_websites_path = r"E:\NewFolder\dabin\福建商业明细表(福建)22.2.10-主席.xlsx"
    set_database_path = r"E:\NewFolder\dabin\产品库-傅镔滢.xlsx"
    set_save_path = r"E:\NewFolder\dabin\中药控股成药营销中心一级商业2024年5月第2周周进销存报表（福建）.xlsx"
    set_week_interval = 1  # 间隔的查询周数
    main(set_chrome_path, set_chromedriver_path, set_websites_path, set_save_path, set_database_path, set_week_interval)
