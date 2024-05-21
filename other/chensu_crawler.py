"""
爬虫抓取苏工作所需的库存资料,并导出为xls表格
XLS属于较老版本,需使用xlwt数据库
接受OCR识别请求地址,  http://localhost:8557/ocr
"""
import os
import xlwt
import datetime
import traceback
import collections
import pandas as pd
import numpy as np
from typing import Tuple, List
from xlwt.Worksheet import Worksheet
from medicine_utils import init_chrome, analyze_website, start_socket, SPFJWeb, TCWeb, DruggcWeb, LYWeb, INCAWeb


class DataToExcel:
    """将数据转化为EXCEL表格类"""

    def __init__(self, save_path, database_path):
        self.save_path = save_path
        self.database = self.init_production_database(database_path)  # 产品数据库
        self.data = collections.defaultdict(lambda: collections.defaultdict(list))
        self.widths = [10, 10, 10, 12, 14]
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

    def init_production_database(self, database_path):
        """获取产品数据库"""
        rd = collections.defaultdict(dict)
        database = pd.read_excel(database_path)
        for _, row in database.iterrows():
            client_production: str = row[1]
            remark = row[3]
            client_production = client_production.replace(" ", "")
            remark = "" if pd.isna(remark) else remark
            rd[row[0]][client_production] = [row[2], remark]
        return rd

    def save(self):
        self.wb.save(self.save_path)

    def len_byte(self, value):
        """获取字符串长度,一个中文的长度为2"""
        value = str(value)
        length = len(value)
        utf8_length = len(value.encode('utf-8'))
        length = (utf8_length - length) / 2 + length
        return int(length) + 2

    def write_to_excel(self, breakpoint_data: pd.DataFrame):
        """将数据写入EXCEL表格"""
        # 写入断点之前的数据
        if breakpoint_data is not None:
            for row in breakpoint_data.itertuples():
                self.write_row(row[1], row[2], row[3], row[7], row[8])
        # 写入抓取的数据
        for client_name, name2standard in self.database.items():
            # 没有抓取这个网站的数据，则跳过
            if client_name not in self.data:
                continue
            web_data = self.data[client_name]
            write_value = []  # 写入的数据:产品标准名称，数量, 参考信息，备注
            for product_name, [standard_name, reference] in name2standard.items():
                number = int(np.sum(web_data.pop(product_name))) if product_name in web_data else 0
                if standard_name == "外用红色诺卡氏菌细胞壁骨架":
                    number = number * 2
                write_value.append([standard_name, number, reference, ""])

            for product_name, number in web_data.items():
                number = int(np.sum(number))
                write_value.append([product_name, number, "", "未在产品库找到"])

            for standard_name, number, reference, remark in write_value:
                self.write_row(client_name, standard_name, number, reference, remark)
            print(f"[{client_name}]已将数据写入到excel表格中.")
        self.save()
        print("所有数据已全部写入完成")

    def cell_format(self):
        """设置单元格格式"""
        for i, width in enumerate(self.widths):
            self.ws.col(i).width = width * 256
        self.save()
        print("格式优化完成")

    def write_row(self, client_name, standard_name, number, reference, remark):
        """写入行数据"""
        self.ws.write(self.row_i, 0, client_name)
        self.ws.write(self.row_i, 1, standard_name)
        self.ws.write(self.row_i, 2, number)
        self.ws.write(self.row_i, 3, self.date, self.date_style)
        self.ws.write(self.row_i, 4, self.date, self.date_style)
        self.ws.write(self.row_i, 6, reference)
        self.ws.write(self.row_i, 7, remark)
        self.row_i += 1

        self.widths[0] = max(self.widths[0], self.len_byte(client_name))
        self.widths[1] = max(self.widths[1], self.len_byte(standard_name))
        self.widths[2] = max(self.widths[2], self.len_byte(str(number)))


def read_breakpoint(path) -> Tuple[set, pd.DataFrame]:
    """读取断点数据"""
    if not os.path.exists(path):
        print("无断点数据,从头到尾抓取")
        return None, None
    datas = pd.read_excel(path)
    client_names = set(datas["一级商业*"].tolist())
    datas.fillna("", inplace=True)
    print(f"检测到断点,进行断点续查。已抓取数据:{client_names}")
    return client_names, datas


def crawler_websites_data(websites_by_code: List[dict], websites_no_code: List[dict], chrome_path, chromedriver_path):

    def crawler_general(datas: List[dict], url2method):
        """抓取的通用方法"""
        for data in datas:
            client_name = data.pop("client_name")
            website_url = data.pop("website_url")
            try:
                method_list = url2method[website_url]
                method_list[0](**data)
                # 不同账号重复商品，只取其中一个账户，除了复方α-酮酸片
                this_account = collections.defaultdict(list)
                for value in method_list[1]():
                    product_name, amount = [value[i] for i in method_list[2]]
                    this_account[product_name].append(amount)
                    if "复方α-酮酸片" in product_name:
                        crawler_data[client_name][product_name].append(amount)
                crawler_data[client_name].update(this_account)

            except Exception:
                print("-" * 150)
                print(f"脚本运行出现异常, 出错的截至问题公司:{client_name}")
                print(traceback.format_exc())
                print("-" * 150)
                print("去除该客户的全部数据")
                crawler_data.pop(client_name)
                return True
        return False

    crawler_data = collections.defaultdict(lambda: collections.defaultdict(list))
    print("抓取无验证码的网站数据")
    driver = init_chrome(chromedriver_path, chrome_path=chrome_path, is_proxy=False)
    spfj = SPFJWeb(driver)
    inca = INCAWeb(driver)
    luyan = LYWeb(driver)
    url_condition = {
        spfj.url: [spfj.login, spfj.get_inventory, [0, 1]],
        inca.url: [inca.login, inca.get_inventory, [0, 1]],
        luyan.url: [luyan.login, luyan.get_inventory, [0, 1]],
    }
    is_error = crawler_general(websites_no_code, url_condition)
    if is_error:
        return crawler_data
    print("关闭浏览器")
    driver.quit()
    print("抓取有验证码的网站数据")
    q = start_socket()
    driver = init_chrome(chromedriver_path, chrome_path=chrome_path)
    tc = TCWeb(driver, q)
    druggc = DruggcWeb(driver, q)
    url_condition = {
        tc.url: [tc.login, tc.get_inventory, [0, 1]],
        druggc.url: [druggc.login, druggc.get_inventory, [0, 1]]
    }
    crawler_general(websites_by_code, url_condition)
    print("关闭浏览器")
    driver.quit()
    return crawler_data


def main(websites_path, chrome_path, chromedriver_path, save_path, database_path):
    print("读取断点数据")
    breakpoint_names, breakpoint_datas = read_breakpoint(save_path)
    print("针对网站数据进行分类")
    websites_by_code, websites_no_code = analyze_website(websites_path, breakpoint_names)
    print("从网站上爬取所需数据")
    crawler_data = crawler_websites_data(websites_by_code, websites_no_code, chrome_path, chromedriver_path)
    print("开始写入所有数据")
    writer = DataToExcel(save_path, database_path)
    writer.data.update(crawler_data)
    writer.write_to_excel(breakpoint_datas)
    print("已完成所有数据写入,开始优化格式")
    writer.cell_format()
    print("程序运行已完成")


if __name__ == "__main__":
    import argparse
    parser = argparse.ArgumentParser()
    parser.add_argument("-p", "--path", type=str, default=r"E:\NewFolder\chensu", help="数据文件的所在文件夹地址")
    opt = {key: value for key, value in parser.parse_args()._get_kwargs()}

    user_folder = opt["path"]
    set_chrome_path = os.path.join(user_folder, r"..\chromedriver_mac_arm64_114\chrome114\App\Chrome-bin\chrome.exe")
    set_chromedriver_path = s.path.join(user_folder, r"..\chromedriver_mac_arm64_114\chromedriver.exe")
    set_websites_path = os.path.join(user_folder, "库存网查明细.xlsx")
    set_database_path = os.path.join(user_folder, "脚本产品库.xlsx")
    now_day = datetime.date.today().strftime('%Y%m%d')
    set_save_path = os.path.join(user_folder, f"库存导入{now_day}.xls")
    main(set_websites_path, set_chrome_path, set_chromedriver_path, set_save_path, set_database_path)
