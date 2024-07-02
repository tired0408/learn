"""
爬虫抓取苏工作所需的库存资料,并导出为xls表格
XLS属于较老版本,需使用xlwt数据库
接受OCR识别请求地址,  http://localhost:8557/ocr
"""
import os
import abc
import copy
import xlwt
import datetime
import traceback
import collections
import pandas as pd
from typing import Tuple, List, Dict
from xlwt.Worksheet import Worksheet
from xlwt.Workbook import Workbook
from dateutil.relativedelta import relativedelta
from medicine_utils import init_chrome, analyze_website, SPFJWeb, TCWeb, DruggcWeb, LYWeb, INCAWeb, CaptchaSocketServer


class Golbal:

    def __init__(self) -> None:
        self.is_deliver = None  # 是否是发货明细
        self.download_path = None  # 下载路径
        self.chrome_path = None  # 谷歌浏览器路径
        self.chromedriver_path = None  # 谷歌浏览器驱动路径
        self.websites_path = None  # 库存网查明细的文件路径
        self.database_path = None  # 脚本产品库的文件路径
        self.save_path = None  # 保存地址
        self.title = None  # 导出文件标题
        self.widths = None  # 导出文件每列的宽度
        self.data = collections.defaultdict(dict)  # 数据

    def set_data(self, path, is_deliver=False):
        """根据选择设置数据"""
        self.is_deliver = is_deliver
        self.download_path = path
        self.chrome_path = os.path.join(path, r"..\chromedriver_mac_arm64_114\chrome114\App\Chrome-bin\chrome.exe")
        self.chromedriver_path = os.path.join(path, r"..\chromedriver_mac_arm64_114\chromedriver.exe")
        self.websites_path = os.path.join(path, "库存网查明细.xlsx")
        self.database_path = os.path.join(path, "脚本产品库.xlsx")
        now_day = datetime.date.today().strftime('%Y%m%d')
        if is_deliver:
            self.save_path = os.path.join(path, f"发货分析表{now_day}.xls")
            self.title = ["一级商业*", "商品信息*", "本期库存*", "库存日期*", "库存获取日期*", "在途", "参考信息",
                          "备注", "近3个月月均销量", "库存周转天数"]
            self.widths = [10, 10, 10, 12, 14, 10, 10, 10, 10, 10]
        else:
            self.save_path = os.path.join(path, f"库存导入{now_day}.xls")
            self.title = ["一级商业*", "商品信息*", "本期库存*", "库存日期*", "库存获取日期*", "在途", "参考信息", "备注"]
            self.widths = [10, 10, 10, 12, 14, 10, 10, 10]

    def get_id(self, client_name, production_name):
        """根据客户名称及商品名称获取唯一ID"""
        return f"{client_name}@@{production_name}"

    def split_id(self, id: str):
        """根据ID分割出客户名称及商品名称"""
        return id.split("@@")


GOL = Golbal()


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

    def write_to_excel(self, breakpoint_data: dict):
        """将数据写入EXCEL表格"""
        # 将网站抓取数据整理成标准数据
        save_data: Dict[str, dict] = {}
        for _, value in GOL.data.items():
            value.pop("客户商品信息")
            save_data[GOL.get_id(value["一级商业*"], value["商品信息*"])] = value
        # 读取断点数据
        if breakpoint_data is not None:
            save_data.update(breakpoint_data)
        # 写入数据
        print("开始写入数据")
        color_style = {
            "red": self.get_color_style("red"),
            "orange": self.get_color_style("orange"),
            "green": self.get_color_style("green"),
        }
        now_date = datetime.datetime.today()
        date_style = xlwt.XFStyle()
        date_style.num_format_str = 'YYYY/MM/DD'
        for row_i, row in enumerate(save_data.values()):
            row_i += 1

            self.ws.write(row_i, 3, now_date, date_style)
            self.ws.write(row_i, 4, now_date, date_style)
            if "库存周转天数" in row:
                day = row.pop("库存周转天数")
                if day > 45:
                    self.ws.write(row_i, GOL.title.index("库存周转天数"), day)
                else:
                    if day <= 15:
                        color = color_style["red"]
                    elif day <= 30:
                        color = color_style["orange"]
                    else:
                        color = color_style["green"]
                    self.ws.write(row_i, GOL.title.index("库存周转天数"), day, color)

            for key, value in row.items():
                col_i = GOL.title.index(key)
                self.ws.write(row_i, col_i, label=value)
                GOL.widths[col_i] = max(GOL.widths[col_i], self.len_byte(value))

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
    def export_inventory(self, client_name) -> dict:
        """获取库存明细"""
        pass

    @abc.abstractmethod
    def export_deliver(self, client_name) -> dict:
        """获取发货明细"""
        pass


class SPFJWebCustom(SPFJWeb, WebAbstract):

    def export_inventory(self, client_name) -> dict:
        rd = collections.defaultdict(lambda: collections.defaultdict(int))
        for product_name, amount, _ in super().get_inventory():
            id = GOL.get_id(client_name, product_name)
            rd[id]["本期库存*"] += amount
        return rd

    def export_deliver(self, client_name) -> dict:
        rd: dict = self.export_inventory(client_name)
        now_date = datetime.datetime.now()
        start_date = (now_date - relativedelta(months=3)).replace(day=1).strftime("%Y-%m-%d")
        end_date = (now_date.replace(day=1) - relativedelta(days=1)).strftime("%Y-%m-%d")
        datas = super().purchase_sale_stock(start_date, end_date)
        for product_name, _, sales, _ in datas:
            id = GOL.get_id(client_name, product_name)
            rd[id]["近3个月月均销量"] += sales
        return calculate_turnover(rd)


class INCAWebCustom(INCAWeb, WebAbstract):

    def export_inventory(self, client_name) -> dict:
        rd = collections.defaultdict(lambda: collections.defaultdict(int))
        inventory_list = super().get_inventory()
        for product_name, amount, _ in inventory_list:
            id = GOL.get_id(client_name, product_name)
            if "外用红色诺卡氏菌细胞壁骨架" in product_name or "胰岛素" in product_name:
                amount = amount * 2
            rd[id]["本期库存*"] += amount
        return rd

    def export_deliver(self, client_name) -> dict:
        rd: dict = self.export_inventory(client_name)
        now_date = datetime.datetime.now()
        start_date = (now_date - relativedelta(months=3)).replace(day=1).strftime("%Y-%m-%d")
        end_date = (now_date.replace(day=1) - relativedelta(days=1)).strftime("%Y-%m-%d")
        sales_list = super().get_sales(start_date, end_date)
        for product_name, amount in sales_list:
            id = GOL.get_id(client_name, product_name)
            if "外用红色诺卡氏菌细胞壁骨架" in product_name or "胰岛素" in product_name:
                amount = amount * 2
            rd[id]["近3个月月均销量"] += amount
        return calculate_turnover(rd)


class LYWebCustom(LYWeb, WebAbstract):

    def export_inventory(self, client_name) -> dict:
        rd = collections.defaultdict(lambda: collections.defaultdict(int))
        inventory_list = super().get_inventory()
        for product_name, amount, _ in inventory_list:
            id = GOL.get_id(client_name, product_name)
            if "外用红色诺卡氏菌细胞壁骨架" in product_name:
                amount = amount * 2
            rd[id]["本期库存*"] += amount
        return rd

    def export_deliver(self, client_name) -> dict:
        rd: dict = self.export_inventory(client_name)
        now_date = datetime.datetime.now()
        start_date = now_date.replace(day=1).replace(month=1).strftime("%Y-%m-%d")
        datas = super().purchase_sale_stock(start_date)
        for product_name, _, sales, _ in datas:
            id = GOL.get_id(client_name, product_name)
            if "外用红色诺卡氏菌细胞壁骨架" in product_name:
                sales = sales * 2
            rd[id]["近3个月月均销量"] += sales
        return calculate_turnover(rd)


class TCWebCustom(TCWeb, WebAbstract):

    def export_inventory(self, client_name) -> dict:
        rd = collections.defaultdict(lambda: collections.defaultdict(int))
        inventory_list = super().get_inventory()
        for product_name, amount in inventory_list:
            id = GOL.get_id(client_name, product_name)
            if "外用红色诺卡氏菌细胞壁骨架" in product_name:
                amount = amount * 2
            rd[id]["本期库存*"] += amount
        return rd

    def export_deliver(self, client_name) -> dict:
        rd: dict = self.export_inventory(client_name)
        now_date = datetime.datetime.now()
        start_date = (now_date - relativedelta(months=3)).replace(day=1).strftime("%Y-%m-%d")
        end_date = (now_date.replace(day=1) - relativedelta(days=1)).strftime("%Y-%m-%d")
        datas = super().get_product_flow(start_date, end_date)
        for product_name, sales in datas:
            id = GOL.get_id(client_name, product_name)
            if "外用红色诺卡氏菌细胞壁骨架" in product_name:
                sales = sales * 2
            rd[id]["近3个月月均销量"] += sales
        return calculate_turnover(rd)


class DruggcWebCustom(DruggcWeb, WebAbstract):

    def export_inventory(self, client_name):
        rd = collections.defaultdict(lambda: collections.defaultdict(int))
        inventory_list = super().get_inventory()
        for product_name, amount, _ in inventory_list:
            id = GOL.get_id(client_name, product_name)
            if "复方α-酮酸片" in product_name and "本期库存*" in GOL.data[id]:
                GOL.data[id]["本期库存*"] += amount
            else:
                rd[id]["本期库存*"] += amount
        return rd

    def export_deliver(self, client_name) -> dict:
        pass


def calculate_turnover(data: dict):
    """计算周转天数"""
    for value in data.values():
        value["近3个月月均销量"] = round(value["近3个月月均销量"] / 3)
        if value["近3个月月均销量"] == 0:
            continue
        value["库存周转天数"] = round(value["本期库存*"] / value["近3个月月均销量"] * 30)
    return data


def read_production_database():
    """获取产品数据库"""
    database = pd.read_excel(GOL.database_path)
    for _, row in database.iterrows():
        client_name = row.iloc[0]
        production_client_name: str = row.iloc[1]
        production_client_name = production_client_name.replace(" ", "")
        production_standard_name: str = row.iloc[2]
        reference = row.iloc[3]
        reference = "" if pd.isna(reference) else reference
        if GOL.is_deliver and client_name in ["厦门片仔癀宏仁医药有限公司", "漳州片仔癀宏仁医药有限公司"]:
            continue
        GOL.data[GOL.get_id(client_name, production_client_name)] = {
            "一级商业*": client_name,
            "商品信息*": production_standard_name,
            "参考信息": reference,
            "客户商品信息": production_client_name,
        }


def read_breakpoint() -> Tuple[set, dict]:
    """读取断点数据"""
    if not os.path.exists(GOL.save_path):
        print("无断点数据,从头到尾抓取")
        return set(), None
    datas = pd.read_excel(GOL.save_path)
    client_names = set(datas["一级商业*"].tolist())
    datas.fillna("", inplace=True)
    print(f"检测到断点,进行断点续查。已抓取数据:{client_names}")
    # 整理断点数据
    rd = {}
    for _, row in datas.iterrows():
        client_name = row["一级商业*"]
        product_name = row["商品信息*"]
        data = row.to_dict()
        data.pop("库存日期*")
        data.pop("库存获取日期*")
        rd[GOL.get_id(client_name, product_name)] = data
    return client_names, rd


def crawler_general(datas: List[dict], url2class: Dict[str, WebAbstract]):
    """抓取的通用方法"""
    for data in datas:
        client_name = data.pop("client_name")
        website_url = data.pop("website_url")
        web_class = url2class[website_url]
        try:
            web_class.login(**data)
            # 不同账号重复商品，只取其中一个账户
            if GOL.is_deliver:
                this_account = web_class.export_deliver(client_name)
            else:
                this_account = web_class.export_inventory(client_name)
            for key, value in this_account.items():
                if key not in GOL.data:
                    client_name, product_name = GOL.split_id(key)
                    value["一级商业*"] = client_name
                    value["商品信息*"] = product_name
                    value["客户商品信息"] = product_name
                    value["备注"] = "未在产品库找到"
                GOL.data[key].update(value)
        except Exception:
            print("-" * 150)
            print(f"脚本运行出现异常, 出错的截至问题公司:{client_name}")
            print(traceback.format_exc())
            print("去除该客户的全部数据")
            data_key = copy.deepcopy(list(GOL.data.keys()))
            for key in data_key:
                if client_name not in key:
                    continue
                GOL.data.pop(key)
            print("-" * 150)
            return True
    return False


def crawler_websites_data(websites_by_code: List[dict], websites_no_code: List[dict]):
    """从网站上抓取数据，并写入全局变量"""
    print("抓取无验证码的网站数据")
    driver = init_chrome(GOL.chromedriver_path, GOL.download_path, chrome_path=GOL.chrome_path, is_proxy=False)
    url2class: Dict[str, WebAbstract] = {
        SPFJWeb.url: SPFJWebCustom(driver),
        INCAWeb.url: INCAWebCustom(driver, GOL.download_path),
        LYWeb.url: LYWebCustom(driver)
    }
    is_error = crawler_general(websites_no_code, url2class)
    if is_error:
        return
    print("关闭浏览器")
    driver.quit()
    print("抓取有验证码的网站数据")
    sock = CaptchaSocketServer()
    driver = init_chrome(GOL.chromedriver_path, GOL.download_path, chrome_path=GOL.chrome_path)
    url2class: Dict[str, WebAbstract] = {
        TCWeb.url: TCWebCustom(driver, sock),
        DruggcWeb.url: DruggcWebCustom(driver, sock, GOL.download_path)
    }
    crawler_general(websites_by_code, url2class)
    print("关闭浏览器")
    driver.quit()


def main(path, is_deliver):
    print("设置全局数据")
    GOL.set_data(path, is_deliver)
    print("读取数据库信息")
    read_production_database()
    print("读取断点数据")
    breakpoint_names, breakpoint_datas = read_breakpoint()
    print("针对网站数据进行分类")
    if is_deliver:
        breakpoint_names.add("厦门片仔癀宏仁医药有限公司")
        breakpoint_names.add("漳州片仔癀宏仁医药有限公司")
    websites_by_code, websites_no_code = analyze_website(GOL.websites_path, breakpoint_names)
    print("从网站上爬取所需数据")
    crawler_websites_data(websites_by_code, websites_no_code)
    print("开始写入所有数据")
    writer = DataToExcel()
    writer.write_to_excel(breakpoint_datas)
    print("程序运行已完成")


if __name__ == "__main__":
    import argparse
    parser = argparse.ArgumentParser()
    parser.add_argument("-p", "--path", type=str, default=r"E:\NewFolder\chensu", help="数据文件的所在文件夹地址")
    parser.add_argument("-d", "--deliver", action="store_true", help="是否导出发货分析表，默认库存导入表")
    opt = {key: value for key, value in parser.parse_args()._get_kwargs()}
    main(opt["path"], opt["deliver"])
