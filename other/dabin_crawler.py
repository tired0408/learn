"""
爬虫抓取大饼所需的进销存报表,并导出为xlsx表格
接受OCR识别请求地址,  http://localhost:8557/ocr
"""
import openpyxl
import traceback
import collections
import pandas as pd
import numpy as np
from typing import Dict, List
from datetime import datetime, timedelta
from openpyxl.cell.cell import Cell, MergedCell
from openpyxl.styles import PatternFill
from openpyxl.utils.cell import get_column_letter
from medicine_utils import init_chrome, analyze_website, SPFJWeb, DruggcWeb, LYWeb, INCAWeb, CaptchaSocketServer


class GolbalData:
    """全局数据"""

    def __init__(self) -> None:
        self.database: collections.defaultdict = None

    def gain_database(self, path):
        """获取数据库"""
        self.database = collections.defaultdict(str)
        database = pd.read_excel(path)
        for _, row in database.iterrows():
            client_production: str = row[1]
            if pd.isna(client_production):
                continue
            client_production = client_production.replace(" ", "")
            self.database[client_production] = row[2]

    def get_standard(self, name):
        """根据商品名获取标准名称"""
        if name not in self.database:
            print(f"############未在产品库里，请及时添加:{name}")
            return None
        return self.database[name]


class WeekData:
    """每个客户周数据类"""

    def __init__(self) -> None:
        self.purchase = collections.defaultdict(list)  # 本周进货
        self.inventory = collections.defaultdict(list)  # 本周库存
        self.code = collections.defaultdict(list)  # 批号
        self.sales = collections.defaultdict(list)  # 本周销量


class MonthData:
    """每个客户的月数据类"""

    def __init__(self) -> None:
        self.purchase = collections.defaultdict(list)  # 本月进货
        self.sales = collections.defaultdict(list)  # 本月销货
        self.code = collections.defaultdict(list)  # 批号


class DataToExcel:
    """读取xlsx文件,并修改里面的数值"""

    def __init__(self, path, interval) -> None:
        self.path = path
        self.wb = openpyxl.load_workbook(path)
        self.ws = self.init_ws()
        self.breakpoint = self.gain_breakpoint(path, "本月销货" if interval == 4 else "本周库存")

    def __del__(self):
        self.wb.close()

    def save(self):
        """保存文件"""
        self.wb.save(self.path)

    def init_ws(self):
        """初始化工作表"""
        ws = self.wb.active
        print("删除颜色")
        for row in ws.iter_rows():
            for cell in row:
                if not isinstance(cell, Cell) and not isinstance(cell, MergedCell):
                    raise Exception("表格存在异常")
                cell.fill = PatternFill(start_color="FFFFFF", end_color="FFFFFF", fill_type="solid")
        return ws

    @staticmethod
    def gain_breakpoint(path, name):
        """读取周数据的断点，获取已经爬取的客户名称"""
        df = pd.read_excel(path, header=1)
        titles = list(df.columns)
        if df.iloc[0, titles.index("统计日期")].strftime("%Y/%m/%d") != datetime.now().strftime("%Y/%m/%d"):
            print("无断点数据，从头开始抓取")
            return
        df = df[df.iloc[:, titles.index("负责人")] == "傅镔滢"]
        df = df.iloc[:, [titles.index("经销商业公司名称"), titles.index(name)]]
        df = df.groupby("经销商业公司名称").agg({"本周库存": "sum"})
        df = df.where(df != 0, np.nan)
        df = df.dropna()
        rd = df.index.to_list()
        print(f"检测到断点,进行断点续查。已抓取数据:{rd}")
        return rd

    def write_week_data(self, all_data: Dict[str, WeekData]):
        """将周数据写入excel表格"""
        now_date = datetime.now()
        titles = [value for row in self.ws.iter_rows(min_row=2, max_row=2, values_only=True) for value in row]
        name_i = titles.index("产品名称、规格")
        person_i = titles.index("负责人")
        company_i = titles.index("经销商业公司名称")
        date_i = titles.index("统计日期")
        last_inventory_i = titles.index("上周库存")
        purchase_i = titles.index("本周进货")
        now_inventory_i = titles.index("本周库存")
        code_i = titles.index("批号")
        remark_i = titles.index("备注")
        for row in self.ws.iter_rows(min_row=3):
            name: Cell = row[name_i]
            person: Cell = row[person_i]
            company: Cell = row[company_i]
            date: Cell = row[date_i]
            last_inventory: Cell = row[last_inventory_i]
            purchase: Cell = row[purchase_i]
            now_inventory: Cell = row[now_inventory_i]
            code: Cell = row[code_i]
            remark: Cell = row[remark_i]
            # 将数据初始化
            if date.value.strftime("%Y/%m/%d") != now_date.strftime("%Y/%m/%d"):
                date.value = now_date
                last_inventory.value = now_inventory.value  # 上周库存
                purchase.value = ""  # 本周进货
                now_inventory.value = ""  # 本周库存
                code.value = ""  # 批号
                remark.value = ""  # 备注(记录破损情况)
            # 填写网站数据
            if person.value == "傅镔滢" and company.value in all_data:
                data = all_data[company.value]
                purchase.value = int(np.sum(list(data.purchase[name.value])))
                now_inventory.value = int(np.sum(list(data.inventory[name.value])))
                code.value = "、".join(list(data.code[name.value])) if name.value in data.code else ""
                sell_index = f"{get_column_letter(titles.index('本周销货') + 1)}{name.row}"
                remark.value = f"=({sell_index}-{int(np.sum(list(data.sales[name.value])))})"


class SPFJWebSelf(SPFJWeb):

    def get_datas(self, user, password, district_name, start_date_str, save_data: WeekData):
        """获取所需数据"""
        self.login(user, password, district_name)
        for product_name, purchase, sale, inventory in self.purchase_sale_stock(start_date_str):
            standard = GOL.get_standard(product_name)
            if standard is None:
                continue
            save_data.purchase[standard].append(purchase)
            save_data.sales[standard].append(sale)
            save_data.inventory[standard].append(inventory)
        for product_name, _, code in self.get_inventory(start_date_str):
            standard = GOL.get_standard(product_name)
            if standard is None:
                continue
            save_data.code[standard].append(code)


class LYWebSelf(LYWeb):

    def get_datas(self, user, password, district_name, start_date_str, save_data: WeekData):
        self.login(user, password, district_name)
        for product_name, _, code in self.get_inventory():
            standard = GOL.get_standard(product_name)
            if standard is None:
                continue
            save_data.code[standard].append(code)
        for product_name, purchase, sale, inventory in self.purchase_sale_stock(start_date_str):
            standard = GOL.get_standard(product_name)
            if standard is None:
                continue
            save_data.purchase[standard].append(purchase)
            save_data.sales[standard].append(sale)
            save_data.inventory[standard].append(inventory)


class DruggcWebSelf(DruggcWeb):

    def get_datas(self, user, password, district_name, start_date_str, save_data: WeekData):
        self.login(user, password, district_name)
        for product_name, inventory, code in self.get_inventory():
            standard = GOL.get_standard(product_name)
            if standard is None:
                continue
            save_data.inventory[standard].append(inventory)
            save_data.code[standard].append(code)
        for product_name, amount in self.get_purchase(start_date_str):
            standard = GOL.get_standard(product_name)
            if standard is None:
                continue
            save_data.purchase[standard].append(amount)
        for product_name, amount in self.get_sales(start_date_str):
            standard = GOL.get_standard(product_name)
            if standard is None:
                continue
            save_data.sales[standard].append(amount)


class INCAWebSelf(INCAWeb):

    def get_datas(self, user, password, start_date_str, save_data: WeekData):
        self.login(user, password)
        for product_name, inventory, code in self.get_inventory():
            standard = GOL.get_standard(product_name)
            if standard is None:
                continue
            save_data.inventory[standard].append(inventory)
            save_data.code[standard].append(code)
        for product_name, amount in self.get_purchase(start_date_str):
            standard = GOL.get_standard(product_name)
            if standard is None:
                continue
            save_data.purchase[standard].append(amount)
        for product_name, amount in self.get_sales(start_date_str):
            standard = GOL.get_standard(product_name)
            if standard is None:
                continue
            save_data.sales[standard].append(amount)


def crawler_from_web(chrome_path, chromedriver_path, websites_by_code, websites_no_code,
                     week_interval) -> Dict[str, WeekData]:
    """从网站上爬取数据"""
    def crawler_general(datas: List[dict], url2method):
        for data in datas:
            client_name = data.pop("client_name")
            website_url = data.pop("website_url")
            try:
                data["start_date_str"] = start_date_str
                data["save_data"] = rd[client_name]
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
    rd = collections.defaultdict(WeekData)
    print("计算开始时间")
    start_date = datetime.now()
    if week_interval == 4:
        start_date = start_date.replace(day=1)
    else:
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
        sock = CaptchaSocketServer()
        driver = init_chrome(chromedriver_path, chrome_path=chrome_path)
        druggc = DruggcWebSelf(driver, sock)
        url_condition = {
            druggc.url: druggc.get_datas
        }
        print("开始抓取")
        crawler_general(websites_by_code, url_condition)
        print("关闭浏览器")
        driver.quit()
    return rd


GOL = GolbalData()


def main(chrome_path, chromedriver_path, websites_path, save_path, database_path, week_interval):
    print("获取数据库")
    GOL.gain_database(database_path)
    print("定义数据写入类")
    writer = DataToExcel(save_path, week_interval)
    print("针对网站数据进行分类")
    websites_by_code, websites_no_code = analyze_website(websites_path, writer.breakpoint)
    print("从网站上爬取数据")
    crawler_data = crawler_from_web(chrome_path, chromedriver_path, websites_by_code, websites_no_code, week_interval)
    print("将数据写入到excel表格中")
    writer.write_week_data(crawler_data)
    writer.save()
    print("程序运行已完成")


if __name__ == "__main__":
    set_chromedriver_path = r'E:\NewFolder\chromedriver_mac_arm64_114\chromedriver.exe'
    set_chrome_path = r"E:\NewFolder\chromedriver_mac_arm64_114\chrome114\App\Chrome-bin\chrome.exe"

    set_websites_path = r"E:\NewFolder\dabin\福建商业明细表(福建)22.2.10-主席.xlsx"
    set_database_path = r"E:\NewFolder\dabin\产品库-傅镔滢.xlsx"
    set_save_path = r"E:\NewFolder\dabin\data.xlsx"
    set_week_interval = 1  # 间隔的查询周数
    main(set_chrome_path, set_chromedriver_path, set_websites_path, set_save_path, set_database_path, set_week_interval)
