"""
爬虫抓取苏工作所需的库存资料,并导出为xls表格
XLS属于较老版本,需使用xlwt数据库
接受OCR识别请求地址,  http://localhost:8557/ocr
"""
import os
import abc
import copy
import xlwt
import time
import datetime
import traceback
import collections
import pandas as pd
from typing import Tuple, List, Dict
from xlwt.Worksheet import Worksheet
from xlwt.Workbook import Workbook
from dateutil.relativedelta import relativedelta
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.action_chains import ActionChains
from medicine_utils import analyze_website, SPFJWeb, TCWeb, DruggcWeb, LYWeb, INCAWeb, CaptchaSocketServer, WEBURL
from crawler_util import select_date_1, init_chrome

class Golbal:

    def __init__(self) -> None:
        self.download_path = None  # 下载路径
        self.chrome_path = None  # 谷歌浏览器路径
        self.chromedriver_path = None  # 谷歌浏览器驱动路径
        self.websites_path = None  # 库存网查明细的文件路径
        self.database_path = None  # 脚本产品库的文件路径
        self.save_path = None  # 保存地址
        self.title = None  # 导出文件标题
        self.widths = None  # 导出文件每列的宽度
        self.save_datas: Dict[str, SaveData] = {}  # 客户名称+客户产品名称
        self.web_datas: Dict[str, WebData] = {}  # 客户名称+产品标准名称

    def set_data(self, path, start_date) -> None:
        """根据选择设置数据"""
        self.download_path = path
        self.last_tidy_date = datetime.datetime.strptime(start_date, '%Y%m%d')
        self.chrome_path = os.path.join(path, r"..\chromedriver_mac_arm64_114\chrome114\App\Chrome-bin\chrome.exe")
        self.chromedriver_path = os.path.join(path, r"..\chromedriver_mac_arm64_114\chromedriver.exe")
        self.websites_path = os.path.join(path, "库存网查明细.xlsx")
        self.database_path = os.path.join(path, "脚本产品库.xlsx")
        self.eas_data_path = os.path.join(path, "EAS发货数据.xls")
        self.last_data_path = os.path.join(path, f"发货分析表{start_date}.xls")
        now_day = datetime.date.today().strftime('%Y%m%d')
        self.save_path = os.path.join(path, f"发货分析表{now_day}.xls")
        self.deliver_path = os.path.join(path, f"发货数据总表{now_day}.xlsx")
        self.title = ["一级商业*", "商品信息*", "本期库存*", "库存日期*", "库存获取日期*", "在途", "备注", "参考信息",
                        "所属账号", "当月销售数量", "近3个月月均销量", "库存周转天数"]
        self.widths = [10, 10, 10, 12, 14, 10, 10, 10, 10, 13, 16, 13]

GOL = Golbal()

class SaveData:
    """需要保存的数据"""
    def __init__(self, client_name, name, now_date, reference) -> None:
        """需要保存到EXCEL中的数据

        Args:
            client_name (str): 客户名称
            name (str): 商品标准名称
            now_date (datetime.datetime): 库存获取日期
            reference (str): 参考信息
        """
        self.first_business = client_name  # 一级商业
        self.production_name = name  # 商品标准名称
        self.inventory = 0  # 本期库存
        self.date = now_date  # 库存日期及库存获取日期
        self.on_road = None  # 在途
        self.remark = None  # 备注
        self.reference = reference  # 参考信息
        self.user = None # 所属账号名称
        self.month_sales = 0  # 当月销售数量
        self.month_sales_average = None  # 近3个月月均销量
        self.inventory_turnover_days = None  # 库存周转天数

    def cal_month_sales_average(self, three_month_sales):
        """计算月均销量"""
        self.month_sales_average = round(three_month_sales / 3) 
        return self.month_sales_average
    
    def cal_turnover_days(self, average, inventory):
        """计算周转天数"""
        if average != 0:
            self.inventory_turnover_days = round(inventory / average * 30)
            return self.inventory_turnover_days
        elif inventory > 0:
            self.inventory_turnover_days = -1
            return -1
        

    def cal_on_road(self, restock, inventory, last_inventory, sales):
        """计算在途数量"""
        amount = restock - inventory + last_inventory - sales
        if amount != 0:
            self.on_road = amount
        return amount

class WebData:
    """网站数据"""
    def __init__(self) -> None:
        self.client_pname = None # 客户的商品名称
        self.conversion_ratio = 1  # 盒支转换比例
        self.inventory = 0  # 本期库存
        self.three_month_sale = 0  # 近3个月销量
        self.month_sale = 0  # 当月销售数量
        self.recent_sale = 0  # 近期销量
        
        self.recent_should_restock = 0  # 近期进货数量(厂家给的发货数量)
        self.last_inventory = 0  # 上期库存
        self.last_on_road = 0  # 上期在途
        
class DataToExcel:
    """将数据转化为EXCEL表格类"""

    def __init__(self):
        self.wb, self.ws = self.init_sheet()

    @staticmethod
    def init_sheet() -> Tuple[Workbook, Worksheet]:
        """初始化单元表"""
        wb = xlwt.Workbook()
        ws: Worksheet = wb.add_sheet("Sheet1")
        for j, value in enumerate(GOL.title):
            ws.write(0, j, value)
        return wb, ws

    def save(self):
        self.wb.save(GOL.save_path)

    def len_byte(self, value):
        """获取字符串长度,一个中文的长度为2"""
        value = str(value)
        length = len(value)
        utf8_length = len(value.encode('utf-8'))
        length = (utf8_length - length) / 2 + length
        return int(length) + 2

    def write_to_excel(self, breakpoint_data: Dict[str, SaveData], restock_datas):
        """将数据写入EXCEL表格"""
        web_datas: Dict[str, WebData] = {}
        # 获取进货数据
        for _, row in restock_datas.iterrows():
            standard_id = get_id(row["客户"], row["商品名称"])
            if standard_id in web_datas:
                web_datas[standard_id].recent_should_restock += 0 if pd.isna(row["数量"]) else row["数量"]
        # 获取上期数据
        last_datas = pd.read_excel(GOL.last_data_path)
        for _, row in last_datas.iterrows():
            standard_id = get_id(row["一级商业*"], row["商品信息*"])
            if standard_id in web_datas:
                web_datas[standard_id].last_inventory = 0 if pd.isna(row["本期库存*"]) else row["本期库存*"]
                web_datas[standard_id].last_on_road = 0 if pd.isna(row["在途"]) else row["在途"]
        save_datas: Dict[str, SaveData] = {}
        for _, data in GOL.save_datas.items():
            standard_id = get_id(data.first_business, data.production_name)
            if standard_id not in web_datas:
                continue
            web_data = web_datas[standard_id]
            data.inventory = web_data.inventory
            data.month_sales = web_data.month_sale
            average = data.cal_month_sales_average(web_data.three_month_sale)
            data.cal_turnover_days(average, web_data.inventory)
            data.cal_on_road(web_data.last_on_road + web_data.recent_should_restock, web_data.inventory, 
                             web_data.last_inventory, web_data.recent_sale)
            save_datas[standard_id] = data
        # 读取断点数据
        if breakpoint_data is not None:
            save_datas.update(breakpoint_data)
        # 写入数据
        print("开始写入数据")
        color_style = {
            "red": self.get_color_style("red"),
            "orange": self.get_color_style("orange"),
            "green": self.get_color_style("green"),
        }
        date_style = xlwt.XFStyle()
        date_style.num_format_str = 'YYYY/MM/DD'
        for row_i, data in enumerate(save_datas.values()):
            row_i += 1
            self.ws.write(row_i, 0, data.first_business)
            self.ws.write(row_i, 1, data.production_name)
            self.ws.write(row_i, 2, data.inventory)
            self.ws.write(row_i, 3, data.date, date_style)
            self.ws.write(row_i, 4, data.date, date_style)
            if data.on_road is not None:
                self.ws.write(row_i, 5, data.on_road)
            if data.reference is not None:
                self.ws.write(row_i, 7, data.reference)
            self.ws.write(row_i, 8, data.user)
            self.ws.write(row_i, 9, data.month_sales)
            self.ws.write(row_i, 10, data.month_sales_average)
            if data.inventory_turnover_days is not None:
                if data.inventory_turnover_days < 0:
                    self.ws.write(row_i, 11, "动销缓慢")
                elif data.inventory_turnover_days <= 15:
                    self.ws.write(row_i, 11, data.inventory_turnover_days, color_style["red"])
                elif data.inventory_turnover_days <= 30:
                    self.ws.write(row_i, 11, data.inventory_turnover_days, color_style["orange"])
                elif data.inventory_turnover_days <= 45:
                    self.ws.write(row_i, 11, data.inventory_turnover_days, color_style["green"])
                else:
                    self.ws.write(row_i, 11, data.inventory_turnover_days)
        # 修改格式
        print("开始修改格式")
        for i, width in enumerate(GOL.widths):
            self.ws.col(i).width = width * 256
        self.save()
        print("所有数据已全部写入完成")        

    def get_color_style(self, color):
        """获取颜色样式"""
        style = xlwt.XFStyle()
        pattern = xlwt.Pattern()
        pattern.pattern = xlwt.Pattern.SOLID_PATTERN
        pattern.pattern_fore_colour = xlwt.Style.colour_map[color]
        style.pattern = pattern
        return style


class WebAbstract(abc.ABC):

    @abc.abstractmethod
    def login(self, *arg, **args):
        pass

    @abc.abstractmethod
    def export_deliver(self, client_name) -> Dict[str, WebData]:
        """获取发货明细"""
        pass


class SPFJWebCustom(SPFJWeb, WebAbstract):

    def export_deliver(self, client_name) -> Dict[str, WebData]:
        rd = collections.defaultdict(WebData)
        now_date = datetime.datetime.now()
        # 获取库存数据,23点更新当天数据
        for product_name, amount, _ in super().get_inventory():
            id = get_id(client_name, product_name)
            rd[id].inventory += amount
        if now_date.hour >= 23:
            datas = now_date.strftime("%Y-%m-%d")
            datas = super().purchase_sale_stock(datas, datas)
            for product_name, _, sales, _ in datas:
                id = get_id(client_name, product_name)
                rd[id].inventory += sales
        # 获取近三个月销量
        start_date = (now_date - relativedelta(months=3)).replace(day=1).strftime("%Y-%m-%d")
        end_date = (now_date.replace(day=1) - relativedelta(days=1)).strftime("%Y-%m-%d")
        datas = super().purchase_sale_stock(start_date, end_date)
        for product_name, _, sales, _ in datas:
            id = get_id(client_name, product_name)
            rd[id].three_month_sale += sales
        # 当月销量
        datas = super().purchase_sale_stock(now_date.replace(day=1).strftime("%Y-%m-%d"), now_date.strftime("%Y-%m-%d"))
        for product_name, _, sales, _ in datas:
            id = get_id(client_name, product_name)
            rd[id].month_sale += sales
        # 获取上一次整理日期至今的销售情况
        start_date = GOL.last_tidy_date.strftime("%Y-%m-%d")
        end_date = (now_date - relativedelta(days=1)).strftime("%Y-%m-%d")
        datas = super().purchase_sale_stock(start_date, end_date)
        for product_name, _, sales, _ in datas:
            id = get_id(client_name, product_name)
            rd[id].recent_sale += sales
        return rd


class INCAWebCustom(INCAWeb, WebAbstract):

    def export_deliver(self, client_name) -> Dict[str, WebData]:
        rd = collections.defaultdict(WebData)
        now_date = datetime.datetime.now()
        # 库存实时更新，需要加上当天销售数据
        inventory_list = super().get_inventory()
        for product_name, amount, _ in inventory_list:
            id = get_id(client_name, product_name)
            rd[id].inventory += amount
        datas = now_date.strftime("%Y-%m-%d")
        datas = super().get_sales(datas, datas)
        for product_name, amount in datas:
            id = get_id(client_name, product_name)
            rd[id].inventory += amount
        
        start_date = (now_date - relativedelta(months=3)).replace(day=1).strftime("%Y-%m-%d")
        end_date = (now_date.replace(day=1) - relativedelta(days=1)).strftime("%Y-%m-%d")
        sales_list = super().get_sales(start_date, end_date)
        for product_name, amount in sales_list:
            id = get_id(client_name, product_name)
            rd[id].three_month_sale += amount

        sales_list = super().get_sales(now_date.replace(day=1).strftime("%Y-%m-%d"), now_date.strftime("%Y-%m-%d"))
        for product_name, amount in sales_list:
            id = get_id(client_name, product_name)
            rd[id].month_sale += amount
        
        start_date = GOL.last_tidy_date.strftime("%Y-%m-%d")
        end_date = (now_date - relativedelta(days=1)).strftime("%Y-%m-%d")
        datas = super().get_sales(start_date, end_date)
        for product_name, sales in datas:
            id = get_id(client_name, product_name)
            rd[id].recent_sale += sales
        return rd


class LYWebCustom(LYWeb, WebAbstract):

    def export_deliver(self, client_name) -> Dict[str, WebData]:
        rd = collections.defaultdict(WebData)
        
        inventory_list = super().get_inventory()
        for product_name, amount, _ in inventory_list:
            id = get_id(client_name, product_name)
            rd[id].inventory += amount
        
        now_date = datetime.datetime.now()
        start_date = now_date.replace(day=1).replace(month=1).strftime("%Y-%m-%d")
        datas = super().purchase_sale_stock(start_date)
        for product_name, _, sales, _ in datas:
            id = get_id(client_name, product_name)
            rd[id].three_month_sale += sales
        
        datas = super().purchase_sale_stock(now_date.replace(day=1).strftime("%Y-%m-%d"), now_date.strftime("%Y-%m-%d"))
        for product_name, _, sales, _ in datas:
            id = get_id(client_name, product_name)
            rd[id].month_sale += sales
        
        start_date = GOL.last_tidy_date.strftime("%Y-%m-%d")
        end_date = (now_date - relativedelta(days=1)).strftime("%Y-%m-%d")
        datas = super().purchase_sale_stock(start_date, end_date)
        for product_name, _, sales, _ in datas:
            id = get_id(client_name, product_name)
            rd[id].recent_sale += sales
        return rd


class TCWebCustom(TCWeb, WebAbstract):

    def export_deliver(self, client_name) -> Dict[str, WebData]:
        rd = collections.defaultdict(WebData)
        now_date = datetime.datetime.now()
        # 数据实时更新，新增当天销售数据
        inventory_list = super().get_inventory()
        for product_name, amount in inventory_list:
            id = get_id(client_name, product_name)
            rd[id].inventory += amount
        datas = now_date.strftime("%Y-%m-%d")
        datas = super().get_product_flow(datas, datas)
        for product_name, amount in datas:
            id = get_id(client_name, product_name)
            rd[id].inventory += amount
        
        start_date = (now_date - relativedelta(months=3)).replace(day=1).strftime("%Y-%m-%d")
        end_date = (now_date.replace(day=1) - relativedelta(days=1)).strftime("%Y-%m-%d")
        datas = super().get_product_flow(start_date, end_date)
        for product_name, sales in datas:
            id = get_id(client_name, product_name)
            rd[id].three_month_sale += sales

        datas = super().get_product_flow(now_date.replace(day=1).strftime("%Y-%m-%d"), now_date.strftime("%Y-%m-%d"))
        for product_name, sales in datas:
            id = get_id(client_name, product_name)
            rd[id].month_sale += sales
        
        start_date = GOL.last_tidy_date.strftime("%Y-%m-%d")
        end_date = (now_date - relativedelta(days=1)).strftime("%Y-%m-%d")
        datas = super().get_product_flow(start_date, end_date)
        for product_name, sales in datas:
            id = get_id(client_name, product_name)
            rd[id].recent_sale += sales
        return rd


class DruggcWebCustom(DruggcWeb, WebAbstract):

    def export_deliver(self, client_name) -> Dict[str, WebData]:
        rd = collections.defaultdict(WebData)

        # TODO 厦门片仔癀的库存数据还未知
        inventory_list = super().get_inventory()
        for product_name, amount, _ in inventory_list:
            id = get_id(client_name, product_name)
            rd[id].inventory += amount

        now_date = datetime.datetime.now()
        d1_end = now_date.replace(day=1) - relativedelta(days=1)
        d1 = d1_end.replace(day=1)
        d2_end = d1 - relativedelta(days=1)
        d2 = d2_end.replace(day=1)
        d3_end = d2 - relativedelta(days=1)
        d3 = d3_end.replace(day=1)
        for start, end in [[d1, d1_end], [d2, d2_end], [d3, d3_end]]:
            datas = super().get_sales(start.strftime("%Y-%m-%d"), end.strftime("%Y-%m-%d"))
            for product_name, sales in datas:
                id = get_id(client_name, product_name)
                rd[id].three_month_sale += sales

        datas = super().get_sales(now_date.replace(day=1).strftime("%Y-%m-%d"), now_date.strftime("%Y-%m-%d"))
        for product_name, sales in datas:
            id = get_id(client_name, product_name)
            rd[id].month_sale += sales
        
        start_date = GOL.last_tidy_date.strftime("%Y-%m-%d")
        end_date = (now_date - relativedelta(days=1)).strftime("%Y-%m-%d")
        datas = super().get_sales(start_date, end_date)
        for product_name, sales in datas:
            id = get_id(client_name, product_name)
            rd[id].recent_sale += sales
        return rd

def get_id(client_name:str, production_name:str):
    """根据客户名称及客户商品名称(或商品标准名称)获取唯一ID"""
    return f"{client_name}@@{production_name}"

def split_id(id:str):
    """根据ID分割出客户名称及客户商品名称"""
    return id.split("@@")

def read_production_database(names):
    """获取产品数据库"""
    database = pd.read_excel(GOL.database_path)
    for _, row in database.iterrows():
        client_name = row["客户名"]
        if client_name not in names:
            continue
        client_production_name: str = row["客户产品名称"]
        client_production_name = client_production_name.replace(" ", "")
        production_standard_name: str = row["产品名称"]
        reference = row["参考信息"]
        reference = "" if pd.isna(reference) else reference
        conversion_ratio = int(row["盒支转换"])

        now_date = datetime.datetime.today()
        client_production_id = get_id(client_name, client_production_name)
        standard_production_id = get_id(client_name, production_standard_name)
        save_data = SaveData(client_name, production_standard_name, now_date, reference)
        web_data = WebData()
        web_data.client_pname = client_production_name
        web_data.conversion_ratio = conversion_ratio
        GOL.save_datas[client_production_id] = save_data
        GOL.web_datas[standard_production_id] = web_data


def read_breakpoint() -> Tuple[set, Dict[str, SaveData]]:
    """读取断点数据"""
    if not os.path.exists(GOL.save_path):
        print("无断点数据,从头到尾抓取")
        return set(), None
    # 整理格式
    raw_data = pd.read_excel(GOL.save_path, engine="xlrd")
    if "库存周转天数" in list(raw_data.columns):
        raw_data["库存周转天数"] = pd.to_numeric(raw_data["库存周转天数"], errors='coerce')
    raw_data = raw_data.fillna("")
    # 获取断点数据
    datas = raw_data.copy()
    datas["本期库存*"] = datas["本期库存*"].fillna(0)
    datas['本期库存*'] = pd.to_numeric(datas['本期库存*'], errors='coerce')
    datas = datas.groupby("一级商业*")["本期库存*"].sum().reset_index()
    datas = datas[datas['本期库存*'] != 0]
    ignore_names = set(datas["一级商业*"].tolist())
    print(f"检测到断点,进行断点续查。已抓取数据:{ignore_names}")
    # 整理断点数据
    rd = {}
    for _, row in raw_data.iterrows():
        client_name = row["一级商业*"]
        if client_name not in ignore_names:
            continue
        if row.isnull().values.any():
            print(row)
        
        data = SaveData(row["一级商业*"], row["商品信息*"], row["库存日期*"], row["参考信息"])
        data.inventory = row["本期库存*"]
        data.on_road = row["在途"] if row["在途"] != "" else None
        data.remark = row["备注"] if row["备注"] != "" else None
        data.month_sales = row["当月销售数量"] 
        data.month_sales_average = row["近3个月月均销量"] 
        data.inventory_turnover_days = row["库存周转天数"] if row["库存周转天数"] != "" else None
    
        rd[get_id(client_name, data.production_name)] = data
    return ignore_names, rd


def crawler_general(datas: List[dict], url2class: Dict[str, WebAbstract]):
    """抓取的通用方法"""
    for data in datas:
        client_name = data.pop("client_name")
        website_url = data.pop("website_url")
        user = data["user"]
        web_class = url2class[website_url]
        try:
            web_class.login(**data)
            # 不同账号重复商品，只取其中一个账户
            this_account: Dict[str, WebData] = web_class.export_deliver(client_name)
            for id, value in this_account.items():
                client_name, client_production_name = split_id(id)
                if id not in GOL.save_datas:
                    save_data_value = SaveData(client_name, client_production_name, datetime.datetime.today(), "未在产品信息库找到")
                    GOL.save_datas[id] = save_data_value
                GOL.save_datas[id].user = user
                standard_name = GOL.save_datas[id].production_name
                standard_id = get_id(client_name, standard_name)
                if standard_id not in GOL.web_datas:
                    value.client_pname = client_production_name
                    value.conversion_ratio = 1
                    GOL.web_datas[id] = value
                    continue
                web_data = GOL.web_datas[standard_id]
                if isinstance(web_class, DruggcWebCustom) and "复方α-酮酸片" in id:
                    web_data.inventory += value.inventory * web_data.conversion_ratio
                    web_data.month_sale += value.month_sale * web_data.conversion_ratio
                    web_data.recent_sale += value.recent_sale * web_data.conversion_ratio
                    web_data.three_month_sale += value.three_month_sale * web_data.conversion_ratio
                else:
                    web_data.inventory = value.inventory * web_data.conversion_ratio
                    web_data.month_sale = value.month_sale * web_data.conversion_ratio
                    web_data.recent_sale = value.recent_sale * web_data.conversion_ratio
                    web_data.three_month_sale = value.three_month_sale * web_data.conversion_ratio
        except Exception:
            print("-" * 150)
            print(f"脚本运行出现异常, 出错的截至问题公司:{client_name},{user},{data['password']}")
            print(traceback.format_exc())
            print("去除该客户的全部数据")
            data_key = copy.deepcopy(list(GOL.web_datas.keys()))
            for key in data_key:
                if client_name not in key:
                    continue
                client_name, standard_name = split_id(key)
                web_data = GOL.web_datas.pop(key)
                GOL.save_datas.pop(get_id(client_name, web_data.client_pname))
            print("-" * 150)
            return True
    return False


def crawler_websites_data(websites_by_code: List[dict], websites_no_code: List[dict]):
    """从网站上抓取数据，并写入全局变量"""
    if len(websites_no_code) != 0:
        print("抓取无验证码的网站数据")
        driver = init_chrome(GOL.chromedriver_path, GOL.download_path, chrome_path=GOL.chrome_path, is_proxy=False)
        url2class: Dict[str, WebAbstract] = {
            WEBURL.spfj: SPFJWebCustom(driver, WEBURL.spfj),
            WEBURL.inca: INCAWebCustom(driver, GOL.download_path, WEBURL.inca),
            WEBURL.ly: LYWebCustom(driver, WEBURL.ly)
        }
        is_error = crawler_general(websites_no_code, url2class)
        if is_error:
            return
        print("关闭浏览器")
        driver.quit()
    if len(websites_by_code) != 0:
        print("抓取有验证码的网站数据")
        sock = CaptchaSocketServer()
        driver = init_chrome(GOL.chromedriver_path, GOL.download_path, chrome_path=GOL.chrome_path)
        url2class: Dict[str, WebAbstract] = {
            WEBURL.xm_tc: TCWebCustom(driver, sock, WEBURL.xm_tc),
            WEBURL.fj_tc: TCWebCustom(driver, sock, WEBURL.fj_tc),
            WEBURL.sm_tc: TCWebCustom(driver, sock, WEBURL.sm_tc),
            WEBURL.druggc: DruggcWebCustom(driver, sock, GOL.download_path, WEBURL.druggc)
        }
        crawler_general(websites_by_code, url2class)
        print("关闭浏览器")
        driver.quit()

def get_deliver_goods(date_value):
    """获取已发出货物的信息"""
    if os.path.exists(GOL.deliver_path):
        rd = pd.read_excel(GOL.deliver_path)
        return rd
    rd = []
    print("从EXCEL中读取金蝶软件的相关数据")
    deliver_data = pd.read_excel(GOL.eas_data_path)
    deliver_data = deliver_data[deliver_data["单据状态"] != "保存"]
    start_date = GOL.last_tidy_date.strftime("%Y-%m-%d")
    end_date = datetime.datetime.now().strftime("%Y-%m-%d")
    deliver_data = deliver_data[(deliver_data["订单日期"] >= start_date) & (deliver_data["订单日期"] < end_date)]
    for _, row in deliver_data.iterrows():
        rd.append({"客户": row["客户"], "商品名称": row["物料名称"], "数量": row["数量"]})
    print("打开浏览器，读取发送给客户的数据")
    for user, passwd in [
        ["18626002881", "Scs@5085618"],
        ["18750776934", "ZGh134679"]
    ]:
        driver = init_chrome(GOL.chromedriver_path, GOL.download_path, chrome_path=GOL.chrome_path, is_proxy=False)
        action = ActionChains(driver)
        # 登录网页
        driver.get("https://i.wanbang.net/home/")
        c1 = EC.visibility_of_element_located((By.ID, "home-user_name"))
        c2 = EC.visibility_of_element_located((By.ID, "account"))
        ele = WebDriverWait(driver, 30).until(EC.any_of(c1, c2))
        if ele.get_attribute("id") == "home-user_name":
            action.move_to_element(ele).perform()
            logout = WebDriverWait(driver, 30).until(EC.visibility_of_element_located((By.XPATH, f"//li[@data-app='logout']")))
            logout.click()
        # 登录用户
        ele = WebDriverWait(driver, 30).until(EC.visibility_of_element_located((By.ID, "account")))
        ele.send_keys(user)
        driver.find_element(By.ID, "pwd").send_keys(passwd)
        driver.find_element(By.CLASS_NAME, "btn-login").click()
        # 进入发货数据查找页面
        now_handle = len(driver.window_handles)
        ele = WebDriverWait(driver, 30).until(EC.visibility_of_element_located((By.XPATH, "//span[text()='协同办公']")))
        ele.click()
        WebDriverWait(driver, 30).until(lambda d: len(d.window_handles) > now_handle)
        driver.switch_to.window(driver.window_handles[-1])
        ele = WebDriverWait(driver, 30).until(EC.visibility_of_element_located((By.XPATH, "//span[text()='我的']")))
        ele.click()
        WebDriverWait(driver, 30).until(EC.visibility_of_element_located((By.ID, "myFlowList-page-title")))
        start_date = date_value
        end_date = datetime.datetime.now()
        end_date = (end_date - relativedelta(days=1))
        pattern = "//div[contains(@class, 'el-date-editor')]/input"
        ele = WebDriverWait(driver, 30).until(EC.element_to_be_clickable((By.XPATH, pattern)))
        ele.click()
        left_container = driver.find_element(By.CLASS_NAME, "is-left")
        right_container = driver.find_element(By.CLASS_NAME, "is-right")
        left_click_btn = left_container.find_element(By.CLASS_NAME, "el-icon-arrow-left")
        left_label_ele = left_container.find_element(By.XPATH, "./div/div")
        right_click_btn = right_container.find_element(By.CLASS_NAME, "el-icon-arrow-right")
        right_label_ele = right_container.find_element(By.XPATH, "./div/div")
        date_pattern = "%Y 年 %m 月"
        day_pattern = [".//span[contains(text(), '", "')]/ancestor::td[@class='available']"]
        select_date_1(start_date, end_date, left_label_ele, right_label_ele, left_container, right_container, date_pattern, day_pattern,
                    left_click_btn, right_click_btn)
        ele = WebDriverWait(driver, 30).until(EC.element_to_be_clickable((By.XPATH, "//button[text()='查询']")))
        ele.click()
        time.sleep(1)
        while True:
            index = 0
            while True:
                data_container = WebDriverWait(driver, 30).until(EC.visibility_of_element_located((By.CLASS_NAME, "cf-flex-content-wrap")))
                if "loading" in data_container.get_attribute("class"):
                    WebDriverWait(data_container, 10).until(lambda d: "loading" not in d.get_attribute("class"))
                table_container = (By.XPATH, ".//div[contains(@class, 'el-table__body-wrapper is-scrolling')]")
                table_container = WebDriverWait(data_container, 10).until(EC.visibility_of_element_located(table_container))
                eles = table_container.find_elements(By.TAG_NAME, "tr")
                if index >= len(eles):
                    break
                ele = eles[index]
                eles = ele.find_elements(By.TAG_NAME, "td")
                ele = eles[3]
                if "产品发货" not in ele.text:
                    index += 1
                    continue
                ele.click()
                detail_container = WebDriverWait(driver, 30).until(EC.visibility_of_element_located((By.ID, "SDK_approval-detail-container")))
                client_name = (By.XPATH, "//span[text()='客户']/following-sibling::span[1]/span/span")
                client_name = WebDriverWait(driver, 30).until(EC.visibility_of_element_located(client_name))
                client_name = client_name.text
                detail_trs = detail_container.find_element(By.CLASS_NAME, "cf-response-table-tbody")
                detail_trs = detail_container.find_elements(By.TAG_NAME, "tr")
                for dtr in detail_trs[1:-1]:
                    detail_tds = dtr.find_elements(By.TAG_NAME, "td")
                    rd.append({"客户": client_name, "商品名称": detail_tds[2].text, "数量": float(detail_tds[4].text)})
                index += 1
                driver.find_element(By.CLASS_NAME, "cf-link-arrow-l").click()
            next_page_btn = driver.find_element(By.CLASS_NAME, "cf-arrow-right")
            if "disabled" in next_page_btn.get_attribute("class"):
                break
            next_page_btn.click()
        print("关闭浏览器")
        driver.quit()
    print("保存发货数据")
    rd = pd.DataFrame(rd)
    rd = rd.sort_values(by=["客户", "商品名称"])
    rd.to_excel(GOL.deliver_path, index=False)
    return rd


def main(path, start_date, ignore_names: List):
    print("设置全局数据")
    GOL.set_data(path, start_date)
    print("获取在途数据")
    restock_datas = get_deliver_goods(GOL.last_tidy_date)
    print("读取断点数据")
    breakpoint_names, breakpoint_datas = read_breakpoint()
    print("针对网站数据进行分类")
    ignore_names.extend(breakpoint_names)
    websites_by_code, websites_no_code, client_names = analyze_website(GOL.websites_path, ignore_names)
    print("读取数据库信息")
    read_production_database(client_names)
    print("从网站上爬取所需数据")
    crawler_websites_data(websites_by_code, websites_no_code)
    print("开始写入所有数据")
    writer = DataToExcel()
    writer.write_to_excel(breakpoint_datas, restock_datas)
    print("程序运行已完成")


if __name__ == "__main__":
    import argparse
    parser = argparse.ArgumentParser()
    parser.add_argument("-p", "--path", type=str, default=r"E:\NewFolder\chensu", help="数据文件的所在文件夹地址")
    parser.add_argument("-d", "--date", type=str, default="20241019", help="上次统计的日期时间")
    parser.add_argument("-t", "--topo", action="store_true", help="是否是局部数据，去掉厦门片仔癀宏仁医药有限公司、漳州片仔癀宏仁医药有限公司")
    opt = {key: value for key, value in parser.parse_args()._get_kwargs()}
    names = ["厦门片仔癀宏仁医药有限公司", "漳州片仔癀宏仁医药有限公司"] if opt["topo"] else []
    main(opt["path"], opt["date"], names)
