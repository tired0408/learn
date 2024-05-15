"""
爬虫抓取夏云所需的销售数据报表,并填入已有的数据表中
"""
import os
import time
import shutil
import warnings
import calendar
import traceback
import xlwings as xw
import pandas as pd
from decimal import Decimal
from abc import ABC, abstractmethod
from selenium.webdriver import Chrome
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.remote.webelement import WebElement
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import NoSuchElementException, TimeoutException
from typing import Dict, List
from datetime import timedelta, date, datetime
from openpyxl import load_workbook
from openpyxl.utils.cell import get_column_letter
warnings.simplefilter("ignore")  # 忽略pandas使用openpyxl读取excel文件的警告


class SavePath:
    """保存地址"""

    def __init__(self) -> None:
        self.operate_detail = None  # 营业明细表
        self.synthesize_operate = None  # 综合营业统计表
        self.synthesize_income = None  # 综合收款统计表
        self.store_consume = None  # 储值消费汇总表
        self.member_addition = None  # 会员新增情况统计表
        self.pay_settlement = None  # 支付结算表

        self.take_out = None  # 外卖收入汇总表
        self.eleme_bill = None  # 饿了么账单明细
        self.autotrophy_meituan = None  # 自营外卖/自提订单明细表(美团)
        self.autotrophy_dada = None  # 自营外卖(达达)


class GolbalData:
    """全局数据"""

    def __init__(self) -> None:
        self.days, self.last_month = self.generate_last_month()
        self.save_path = SavePath()

    def generate_last_month(self):
        """获取上个月所有日期"""
        today = date.today()
        first_day_of_this_month = today.replace(day=1)
        last_day_of_last_month = first_day_of_this_month - timedelta(days=1)
        year, month = last_day_of_last_month.year, last_day_of_last_month.month
        num_days = calendar.monthrange(year, month)[1]
        last_months = [date(year, month, day).strftime("%y.%m.%d") for day in range(1, num_days + 1)]
        return last_months, last_day_of_last_month


GOL = GolbalData()


class PerDayData:
    """营业明细统计表的每日数据类"""

    def __init__(self) -> None:
        self.cash = None  # 现金
        self.wechat = None  # 第三方收入（微信）
        self.eat_in = None  # 堂食扫码收入
        self.ele_me = None  # 饿了么收入
        self.member_cash = []  # 会员充值现金收入
        self.member_wechat = []  # 会员充值第三方支付（微信）
        self.member_scan = []  # 会员扫码充值
        self.settlement_amount = None  # 实际结算金额
        self.main_consume = None  # 会员主账号消费
        self.gift_consume = None  # 会员赠送账号消费
        self.main_paid = None  # 会员主账号充值
        self.gift_paid = None  # 会员赠送账号充值
        self.new_member = None  # 新开卡会员
        self.ele_me_free = None  # 饿了么优免金额
        self.other_free = None  # 其他优免金额
        self.hand_charge = None  # 手续费
        self.pubilc_relation_paid = []  # 公关充值金额
        self.pubilc_relation_income = None  # 公关收入


class GetOperateDetail:
    """营业明细表的数据获取类"""

    def __init__(self, path) -> None:
        self.save_data = self.init_save_data()
        self.wb = self.init_excel(path)
        self.ws = self.wb.active

    def init_excel(self, path):
        """调整营业明细表的模板"""
        days_len = len(GOL.days)
        if days_len == 31:
            return load_workbook(path)
        with xw.App(visible=False) as app:
            with app.books.open(path) as wb:
                ws = wb.sheets["营业月报"]
                for i in range(31 - days_len):
                    del_row = 35 - i
                    ws.range(f'{del_row}:{del_row}').api.EntireRow.Delete(Shift=-4162)
                wb.save(GOL.save_path.operate_detail)
        return load_workbook(GOL.save_path.operate_detail)

    def init_save_data(self) -> Dict[str, PerDayData]:
        """初始化每天的保存数据"""
        return {key: PerDayData() for key in GOL.days}

    def write_and_save(self):
        # 设置开始、结束时间
        self.ws.cell(2, 2, GOL.days[0])
        self.ws.cell(3, 2, GOL.days[-1])
        # 设置每天的数据
        for index, day in enumerate(GOL.days):
            row_index = 5 + index
            day_data = self.save_data[day]
            self.ws.cell(row_index, 2, day)
            self.ws.cell(row_index, 3, day_data.cash)
            self.ws.cell(row_index, 4, day_data.wechat)
            self.ws.cell(row_index, 5, day_data.eat_in)
            self.ws.cell(row_index, 7, day_data.ele_me)
            self.ws.cell(row_index, 8, day_data.member_cash)
            self.ws.cell(row_index, 9, day_data.member_wechat)
            self.ws.cell(row_index, 10, day_data.member_scan)
            self.ws.cell(row_index, 13, day_data.settlement_amount)
            self.ws.cell(row_index, 15, day_data.main_consume)
            self.ws.cell(row_index, 16, day_data.gift_consume)
            self.ws.cell(row_index, 17, day_data.main_paid)
            self.ws.cell(row_index, 18, day_data.gift_paid)
            self.ws.cell(row_index, 19, day_data.new_member)
            self.ws.cell(row_index, 24, day_data.ele_me_free)
            self.ws.cell(row_index, 25, day_data.other_free)
            self.ws.cell(row_index, 31, day_data.hand_charge)
            self.ws.cell(row_index, 32, day_data.pubilc_relation_paid)
            self.ws.cell(row_index, 33, day_data.pubilc_relation_income)
        # 保存文件
        self.wb.save(GOL.save_path.operate_detail)

    def read_general_business(self):
        """读取综合营业统计表的相关数据"""
        data = pd.read_excel(GOL.save_path.synthesize_operate, header=None)
        change_i = 0
        assert data.iloc[2, 0] == "营业日期", "表格发生变化，请联系管理员"
        assert data.iloc[2, 11] == "渠道营业构成", "表格发生变化，请联系管理员"
        assert data.iloc[3, 23] == "饿了么外卖", "表格发生变化，请联系管理员"
        assert data.iloc[4, 24] == "营业收入（元）", "表格发生变化，请联系管理员"
        ele_me_i = 25
        assert data.iloc[2, 42] == "营业收入构成", "表格发生变化，请联系管理员"
        assert data.iloc[3, 42] == "现金", "表格发生变化，请联系管理员"
        assert data.iloc[4, 42] == "人民币", "表格发生变化，请联系管理员"
        cach_i = 43
        assert data.iloc[3, 43] == "扫码支付", "表格发生变化，请联系管理员"
        assert data.iloc[4, 43] == "微信", "表格发生变化，请联系管理员"
        assert data.iloc[4, 44] == "支付宝", "表格发生变化，请联系管理员"
        assert data.iloc[4, 45] == "银联二维码（信用卡）", "表格发生变化，请联系管理员"
        # 2024年3月份的表格与2024年1月份的表格存在不同之处
        if data.iloc[4, 46] == "卡余额消费-储值余额":
            eat_in_i_list = [44, 45, 46]
        elif data.iloc[4, 46] == "银联二维码（储蓄卡）":
            eat_in_i_list = [44, 45, 46, 47]
            change_i += 1
        else:
            raise Exception("表格发生变化，请联系管理员")
        assert data.iloc[3, 47 + change_i] == "自定义记账", "表格发生变化，请联系管理员"
        if data.iloc[4, 47 + change_i] == "公关/奖品/活动/无实质性收入（自）":
            pubilc_relation_income_i = 47 + change_i + 1
        elif data.iloc[4, 47 + change_i] == "微信收款（店长号收款）（自）":
            pubilc_relation_income_i = None
            change_i -= 1
        else:
            raise Exception("表格发生变化，请联系管理员")
        assert data.iloc[4, 48 + change_i] == "微信收款（店长号收款）（自）", "表格发生变化，请联系管理员"
        wechat_i = 48 + change_i + 1
        assert data.iloc[2, 51 + change_i] == "支付优惠构成", "表格发生变化，请联系管理员"
        assert data.iloc[3, 52 + change_i] == "外卖", "表格发生变化，请联系管理员"
        assert data.iloc[4, 52 + change_i] == "饿了么外卖", "表格发生变化，请联系管理员"
        ele_me_free_i = 52 + change_i + 1
        assert data.iloc[2].dropna().iloc[-1] == "折扣优惠构成", "表格发生变化，请联系管理员"
        assert data.iloc[3, -1] == "小计", "表格发生变化，请联系管理员"
        other_free_i = -1
        for row in data.iloc[5:].itertuples():
            day_str = row[1]
            if day_str == "合计":
                break
            day_str: str = day_str[2:]
            day_str = day_str.replace("/", ".")
            day_data = self.save_data[day_str]
            day_data.cash = row[cach_i]
            day_data.wechat = row[wechat_i]
            day_data.eat_in = sum([row[i] for i in eat_in_i_list])
            day_data.ele_me = row[ele_me_i]
            day_data.ele_me_free = row[ele_me_free_i]
            day_data.other_free = row[other_free_i]
            day_data.pubilc_relation_income = 0 if pubilc_relation_income_i is None else row[pubilc_relation_income_i]

    def read_general_collection(self):
        """读取综合收款统计表的相关数据"""
        data = pd.read_excel(GOL.save_path.synthesize_income, header=None)
        change_i = 0
        assert data.iloc[2, 0] == "营业日期", "表格发生变化，请联系管理员"
        assert data.iloc[2, 1] == "业务大类", "表格发生变化，请联系管理员"
        assert data.iloc[2, 2] == "业务小类", "表格发生变化，请联系管理员"
        assert data.iloc[2, 4] == "现金", "表格发生变化，请联系管理员"
        assert data.iloc[3, 4] == "人民币", "表格发生变化，请联系管理员"
        assert data.iloc[2, 5] == "扫码支付", "表格发生变化，请联系管理员"
        assert data.iloc[3, 5] == "微信", "表格发生变化，请联系管理员"
        assert data.iloc[3, 6] == "支付宝", "表格发生变化，请联系管理员"
        assert data.iloc[3, 7] == "银联二维码（信用卡）", "表格发生变化，请联系管理员"
        if data.iloc[3, 8] == "卡余额消费-储值余额":
            scan_i_list = [6, 7, 8]
        elif data.iloc[3, 8] == "银联二维码（储蓄卡）":
            scan_i_list = [6, 7, 8, 9]
            change_i += 1
        else:
            raise Exception("表格发生变化，请联系管理员")
        assert data.iloc[2, 9 + change_i] == "自定义记账", "表格发生变化，请联系管理员"
        if data.iloc[3, 9 + change_i] == "公关/奖品/活动/无实质性收入（自）":
            income_i = 9 + change_i + 1
        elif data.iloc[3, 9 + change_i] == "微信收款（店长号收款）（自）":
            income_i = None
            change_i -= 1
        else:
            raise Exception("表格发生变化，请联系管理员")
        assert data.iloc[3, 10 + change_i] == "微信收款（店长号收款）（自）", "表格发生变化，请联系管理员"
        wechat_i = 10 + change_i + 1
        for row in data.iloc[5:].itertuples():
            day_str = row[1]
            if day_str == "合计":
                break
            day_str: str = day_str[2:]
            day_str = day_str.replace("/", ".")
            day_data = self.save_data[day_str]
            if row[2] != "会员充值":
                continue
            if row[3] in ["充值", "撤销充值"]:
                cash = row[5]
                wechat = row[wechat_i]
                scan = row[6] + row[7]
                scan = sum([row[i] for i in scan_i_list])
                income = 0 if income_i is None else row[income_i]
            elif row[3] == "退卡":
                cash = -row[5]
                assert cash <= 0, "退卡金额应该小于0"
                wechat = -row[wechat_i]
                assert wechat <= 0, "退卡金额应该小于0"
                scan = 0
                for i in scan_i_list:
                    assert row[i] >= 0, "退卡金额应该小于0"
                    scan -= row[i]
                income = 0 if income_i is None else -row[income_i]
                assert income <= 0, "退卡金额应该小于0"
            else:
                continue
            day_data.member_cash.append(cash)
            day_data.member_wechat.append(wechat)
            day_data.member_scan.append(scan)
            day_data.pubilc_relation_paid.append(income)
        # 整理数据
        for _, day_data in self.save_data.items():
            assert len(day_data.member_cash) <= 3, "会员充值、退卡数据各自最多只有一条"
            day_data.member_cash = sum(day_data.member_cash)
            assert len(day_data.member_wechat) <= 3, "会员充值、退卡数据各自最多只有一条"
            day_data.member_wechat = sum(day_data.member_wechat)
            assert len(day_data.member_scan) <= 3, "会员充值、退卡数据各自最多只有一条"
            day_data.member_scan = sum(day_data.member_scan)
            assert len(day_data.pubilc_relation_paid) <= 3, "公关/奖品/活动/无实质性收入的充值、退卡数据各自最多只有一条"
            day_data.pubilc_relation_paid = sum(day_data.pubilc_relation_paid)

    def read_store_consume(self):
        """读取储值消费汇总表的相关数据"""
        data = pd.read_excel(GOL.save_path.store_consume, header=None)
        assert data.iloc[2, 0] == "日期", "表格发生变化，请联系管理员"
        assert data.iloc[2, 1] == "储值合计", "表格发生变化，请联系管理员"
        assert data.iloc[3, 1] == "储值余额", "表格发生变化，请联系管理员"
        assert data.iloc[3, 2] == "赠送余额", "表格发生变化，请联系管理员"
        assert data.iloc[2, 4] == "消费合计", "表格发生变化，请联系管理员"
        assert data.iloc[3, 4] == "储值余额", "表格发生变化，请联系管理员"
        assert data.iloc[3, 5] == "赠送余额", "表格发生变化，请联系管理员"
        for row in data.iloc[4:].itertuples():
            day_str = row[1]
            if day_str == "合计":
                break
            day_str: str = day_str[2:]
            day_str = day_str.replace("-", ".")
            day_data = self.save_data[day_str]
            day_data.main_consume = row[5]
            day_data.gift_consume = row[6]
            day_data.main_paid = row[2]
            day_data.gift_paid = row[3]

    def read_newly_increased(self):
        """"读取会员新增情况统计表的相关数据"""
        data = pd.read_excel(GOL.save_path.member_addition, header=None)
        assert data.iloc[2, 0] == "日期", "表格发生变化，请联系管理员"
        assert data.iloc[2, 1] == "合计", "表格发生变化，请联系管理员"
        for row in data.iloc[3:].itertuples():
            if row[1] == "合计":
                break
            day_str: str = row[1][2:]
            day_str = day_str.replace("-", ".")
            day_data = self.save_data[day_str]
            day_data.new_member = row[2]

    def read_pay_settlement(self):
        """读取支付结算表的相关数据"""
        data = pd.read_excel(GOL.save_path.pay_settlement, header=None)
        assert data.iloc[2, 2] == "结算日期", "表格发生变化，请联系管理员"
        assert data.iloc[2, 3] == "交易金额(元)", "表格发生变化，请联系管理员"
        assert data.iloc[2, 5] == "手续费(元)", "表格发生变化，请联系管理员"
        for row in data.iloc[3:len(GOL.days) + 3].itertuples():
            day_str = row[3]
            if pd.isna(day_str):
                break
            day_str: str = day_str[2:]
            day_str = day_str.replace("-", ".")
            day_data = self.save_data[day_str]
            day_data.hand_charge = row[6]
            day_data.settlement_amount = row[4]


class WebCrawler(ABC):

    def __init__(self, driver: Chrome, download_path) -> None:
        self._download_timeout = 60
        self._download_path = download_path
        self._download_file = {}
        self._name2save = self._init_name2save()
        self._driver = driver
        self._action = ActionChains(self._driver)

    @abstractmethod
    def _init_name2save(self):
        """初始化下载文件与文件保存地址的对应关系"""

    def wait_download_finnish(self):
        # 等到下载完成
        for _, file_name in self._download_file.items():
            st = time.time()
            while True:
                if (time.time() - st) > self._download_timeout:
                    raise Exception("Waiting download timeout.")
                if os.path.exists(os.path.join(self._download_path, file_name)):
                    break
        # 移动下载文件
        for key, file_name in self._download_file.items():
            src = os.path.join(self._download_path, file_name)
            shutil.move(src, self._name2save[key])


class DadaCrawler(WebCrawler):

    def _init_name2save(self):
        return {
            "门店订单明细": GOL.save_path.autotrophy_dada
        }

    def login(self):
        """登录达达网站"""
        self._driver.get(r"https://newopen.imdada.cn/#/manager/shop/report/order")

    def download_store_report(self):
        """下载门店报表"""
        # 日期选择
        pattern = (By.XPATH, ".//input[@placeholder='请选择时间']")
        WebDriverWait(self._driver, 10).until(EC.element_to_be_clickable(pattern))
        start_select_ele, end_select_ele = self._driver.find_elements(*pattern)
        start_ele, end_ele = self._driver.find_elements(By.CLASS_NAME, "datepicker")
        self.__date_selection(start_select_ele, start_ele, 1)
        self.__date_selection(end_select_ele, end_ele, GOL.days[-1].split(".")[-1])
        self._driver.find_element(By.XPATH, "//span[text()='搜索']/..").click()
        load_icon = self._driver.find_element(By.CLASS_NAME, "loading-mask")
        WebDriverWait(self._driver, 10).until(EC.visibility_of(load_icon))
        WebDriverWait(self._driver, 10).until(EC.invisibility_of_element(load_icon))
        # 申请报表下载
        btn_str = "//div[text()='Still bread 还是面包厨房（华瑞花园1期店）']/../..//span[text()='下载门店订单明细']"
        self._driver.find_element(By.XPATH, btn_str).click()
        ele = WebDriverWait(self._driver, 10).until(EC.visibility_of_element_located((By.CLASS_NAME, "modal-content")))
        ele = WebDriverWait(ele, 10).until(EC.element_to_be_clickable((By.CLASS_NAME, "close")))
        ele.click()
        # 跳转到下载页面
        ele = WebDriverWait(self._driver, 1).until(EC.element_to_be_clickable((By.XPATH, "//div[text()=' 订单报表']/..")))
        if "active" not in ele.get_attribute("class"):
            ele.click()
        ele = WebDriverWait(ele, 1).until(EC.element_to_be_clickable((By.XPATH, ".//a[text()='下载列表']/..")))
        ele.click()
        WebDriverWait(self._driver, 10).until(EC.visibility_of(load_icon))
        WebDriverWait(self._driver, 10).until(EC.invisibility_of_element(load_icon))
        # 定位到所需要下载的那一行
        now_str = date.today().strftime("%Y-%m-%d")
        true_content = f"(20{GOL.days[0].replace('.', '-')} ~ 20{GOL.days[-1].replace('.', '-')})"
        tr_ele_list = self._driver.find_elements(By.XPATH, f"//div[text()='{now_str}']/../..")
        for tr_ele in tr_ele_list:
            content_ele = tr_ele.find_elements(By.TAG_NAME, "td")[2]
            content_ele = content_ele.find_elements(By.XPATH, "./div/div/div")[1]
            if content_ele.text != true_content:
                continue
            break
        else:
            raise Exception("找不到对应的下载信息")
        # 下载文件
        button = WebDriverWait(tr_ele, 1).until(EC.element_to_be_clickable((By.XPATH, ".//a[text()='下载']")))
        button.click()
        name = os.path.basename(button.get_attribute("href"))
        self._download_file["门店订单明细"] = name
        self.__wait_download(name)

    def __wait_download(self, name):
        """等待开始下载"""
        st = time.time()
        while True:
            if (time.time() - st) > self._download_timeout:
                raise Exception("Waiting download timeout.")
            path = os.path.join(self._download_path, name)
            if os.path.exists(path):
                break
            if os.path.exists(f"{path}.crdownload"):
                break
        print(f"文件已经在下载:{name}")

    def __date_selection(self, select_ele: WebElement, element: WebElement, value):
        """日期选择"""
        select_ele.click()
        WebDriverWait(self._driver, 1).until(EC.visibility_of(element))
        now_date_ele = element.find_element(By.CLASS_NAME, "datepicker-caption")
        need_date_str = GOL.last_month.strftime("%Y年%m月")
        while True:
            cur_date_str = now_date_ele.text
            if cur_date_str == need_date_str:
                break
            now_date = datetime.strptime(cur_date_str, "%Y年%m月").date()
            if now_date - GOL.last_month > timedelta(0):
                element.find_element(By.CLASS_NAME, "dada-ico-angle-left").click()
            else:
                element.find_element(By.CLASS_NAME, "dada-ico-angle-right").click()
            WebDriverWait(now_date_ele, 1).until(lambda ele: ele.text != cur_date_str)
        element.find_element(By.XPATH, f".//div[@class='datepicker-item-text' and text()='{value}']").click()
        WebDriverWait(self._driver, 1).until(EC.invisibility_of_element(element))


class MeiTuanCrawler(WebCrawler):
    """美团网站抓取类"""

    def _init_name2save(self):
        return {
            "综合营业统计": GOL.save_path.synthesize_operate,
            "综合收款统计": GOL.save_path.synthesize_income,
            "储值消费汇总表": GOL.save_path.store_consume,
            "会员新增情况统计表": GOL.save_path.member_addition,
            "支付结算": GOL.save_path.pay_settlement,
            "自营外卖_自提订单明细": GOL.save_path.autotrophy_meituan,
        }

    def login(self):
        """登录显示非法请求，不让操作"""
        # driver.get(r"https://pos.meituan.com")
        # wait.until(EC.visibility_of_element_located((By.CLASS_NAME, "login-section-36Fr9")))
        # driver.switch_to.frame(driver.find_element(By.XPATH, "//iframe[@title='rms epassport']"))
        # wait.until(EC.visibility_of_element_located((By.XPATH, "//a[text()='帐号登录']")))
        # driver.find_element(By.XPATH, "//a[text()='帐号登录']").click()
        # wait.until(EC.visibility_of_element_located((By.ID, "password")))
        # clear_and_send(driver.find_element(By.ID, "account"), r"13194080718")
        # clear_and_send(driver.find_element(By.ID, "password"), r"stillrelax0601")
        # driver.find_element(By.XPATH, "//button[text()='登录']").click()
        self._driver.get(r"https://pos.meituan.com")

    def download_synthesize_operate(self):
        """下载综合营业统计表"""
        menu_name, name = "营业报表", "综合营业统计"
        module = self.__enter_main_module("报表中心")
        submodule = self.__enter_rc_module(module, menu_name, name)
        self.__date_select_1(submodule)
        self.__search(submodule)
        self.__clear_excel(name)
        self.__wait_search_finnsh_1(submodule)
        self.__download_direct(submodule, name)

    def download_autotrophy(self):
        """下载自营外卖/自提订单明细表"""
        menu_name, name = "营业报表", "自营外卖/自提订单明细"
        download_name = name.replace("/", "_")
        module = self.__enter_main_module("报表中心")
        submodule = self.__enter_rc_module(module, menu_name, name)
        self.__date_select_3(submodule)
        self.__search(submodule)
        self.__clear_excel(download_name)
        self.__wait_search_finnsh_1(submodule)
        self.__download_autotrophy_detail(submodule, download_name)

    def download_synthesize_income(self):
        """下载综合收款统计表"""
        menu_name, name = "收款报表", "综合收款统计"
        module = self.__enter_main_module("报表中心")
        submodule = self.__enter_rc_module(module, menu_name, name)
        self.__date_select_1(submodule)
        self.__synthesize_income_condition(submodule)
        self.__search(submodule)
        self.__clear_excel(name)
        self.__wait_search_finnsh_1(submodule)
        self.__download_direct(submodule, name)

    def download_pay_settlement(self):
        """下载支付结算表"""
        menu_name, name = "收款报表", "支付结算"
        module = self.__enter_main_module("报表中心")
        submodule = self.__enter_rc_module(module, menu_name, name)
        self.__date_select_1(submodule)
        self.__pay_settlement_condition(submodule)
        self.__search(submodule)
        self.__clear_excel(name)
        self.__wait_search_finnsh_1(submodule)
        self.__download_direct(submodule, name)

    def download_store_consume(self):
        """下载储值消费汇总表"""
        menu_name, name = "数据报表", "储值消费汇总表"
        module = self.__enter_main_module("营销中心")
        submodule = self.__enter_mc_module(module, menu_name, name)
        self.__date_select_2(submodule)
        self.__search(submodule)
        self.__clear_excel(name)
        self.__wait_search_finnsh_2(submodule)
        self.__download_direct(submodule, name)

    def download_member_addition(self):
        """下载会员新增情况统计表"""
        menu_name, name = "用户", "会员新增情况统计表"
        module = self.__enter_main_module("营销中心")
        submodule = self.__enter_mc_module(module, menu_name, name)
        self.__toggle_old()
        self.__date_select_3(submodule)
        self.__search(submodule)
        self.__clear_excel(name)
        self.__wait_search_finnsh_1(submodule)
        self.__download_direct(submodule, name)

    def __enter_main_module(self, name):
        """进入主模块的方法：运营中心、营销中心、库存管理、报表中心"""
        self._driver.switch_to.default_content()
        # 获取当前容器显示的模块内容
        pattern = (By.XPATH, "//div[@class='main-app']/div[@style='display: block;'][./*]")
        WebDriverWait(self._driver, 10).until(EC.presence_of_all_elements_located(pattern))
        current_module = self._driver.find_element(*pattern)
        # 进入所定模块
        pattern = (By.XPATH, f"//header//span[text()='{name}']/..")
        WebDriverWait(self._driver, 10).until(EC.visibility_of_element_located(pattern))
        element = self._driver.find_element(*pattern)
        if "active-first-menu" in element.get_attribute("class"):
            print(f"已在{name}模块")
            return current_module
        element.click()
        # 等待原有模块消失，即新模块显示
        WebDriverWait(current_module, 10).until(lambda ele: ele.get_attribute("style") == "display: none;")
        # 获取新模块
        pattern = (By.XPATH, "//div[@class='main-app']/div[@style='display: block;'][./*]")
        current_module = self._driver.find_element(*pattern)
        try:
            current_module.find_element(By.XPATH, f".//div[@role='tablist']//span[text()='{name}首页']")
        except NoSuchElementException:
            raise Exception("进入模块出现问题")
        return current_module

    def __enter_rc_module(self, module, menu_name, name):
        """进入报表中心的子模块"""
        submodule = self.__get_statement_submodule(module)
        if self.__get_active_submodule_name(module) != name:
            self.__hover_and_click(module, menu_name, name)
            WebDriverWait(submodule, 10).until(lambda ele: ele.get_attribute("style") == "display: none;")
            submodule = self.__get_statement_submodule(module)
        return submodule

    def __get_statement_submodule(self, module: WebElement):
        """获取报表中心激活的子模块"""
        pattern = (By.XPATH, ".//div[@id='__root_wrapper_rms-report']//div[@style='display: block;']")
        WebDriverWait(module, 10).until(EC.presence_of_all_elements_located(pattern))
        current_submodule = module.find_element(*pattern)
        return current_submodule

    def __enter_mc_module(self, module: WebElement, menu_name, name):
        """进入营销中心的子模块"""
        if self.__get_active_submodule_name(module) != name:
            self.__hover_and_click(module, menu_name, name)
        name2url = {
            "储值消费汇总表": '/web/crm-smart/report/summary-store',
            "会员新增情况统计表": '/web/member/statistic/member-increase#/'
        }
        pattern = (By.XPATH, f".//iframe[@data-current-url='{name2url[name]}']")
        WebDriverWait(module, 60).until(EC.visibility_of_element_located(pattern))
        iframe = module.find_element(*pattern)
        self._driver.switch_to.frame(iframe)

    def __get_active_submodule_name(self, module: WebElement):
        """获取激活的子模块名字"""
        pattern = (By.XPATH, ".//div[@role='tablist']//div[@aria-selected='true']")
        cur_submodule_name = module.find_element(*pattern).text
        return cur_submodule_name

    def __hover_and_click(self, module: WebElement, menu_name, name):
        """悬停菜单栏并点击子模块"""
        pattern = (By.XPATH, f".//div[@class='menu-container ']//span[text()='{menu_name}']/../..")
        WebDriverWait(module, 10).until(EC.visibility_of_element_located(pattern))
        menu = module.find_element(*pattern)
        menu_id = menu.get_attribute("id").split("_")[1]
        for _ in range(4):
            try:
                self._action.move_to_element(menu).perform()
                pattern = (By.XPATH, f"//ul[@id='{menu_id}$Menu']//li[text()='{name}']")
                WebDriverWait(self._driver, 1).until(EC.visibility_of_element_located(pattern))
                self._driver.find_element(*pattern).click()
            except TimeoutException:
                continue
            break

    def __date_select_1(self, submodule: WebElement):
        """日期选择1:综合收款统计的那个类型日期选择控件"""
        pattern = (By.XPATH, ".//input[@placeholder='请选择日期']")
        WebDriverWait(submodule, 10).until(EC.visibility_of_element_located(pattern))
        submodule.find_element(*pattern).click()
        pattern = (By.XPATH, "//div[@class='ant-calendar-footer']//span[text()='上月']")
        WebDriverWait(self._driver, 10).until(EC.visibility_of_element_located(pattern))
        self._driver.find_element(*pattern).click()
        WebDriverWait(self._driver, 10).until_not(EC.visibility_of_element_located(pattern))

    def __date_select_2(self, submodule: WebElement):
        """日期选择2:储值消费汇总表的那个类型日期选择控件"""
        submodule.find_element(By.CLASS_NAME, "el-range-input").click()
        pattern = (By.XPATH, ".//div[@class='el-picker-panel el-date-range-picker el-popper']")
        WebDriverWait(self._driver, 10).until(EC.visibility_of_element_located(pattern))
        left, right = self._driver.find_element(*pattern).find_elements(By.XPATH, "./*/*/*")
        now_date = datetime.today()
        while True:
            try:
                right.find_element(By.CLASS_NAME, "today")
                break
            except NoSuchElementException:
                pass
            right_month_ele = right.find_element(By.XPATH, "./div/div")
            right_month_str = right_month_ele.text
            right_month = datetime.strptime(right_month_str, '%Y 年 %m 月')
            if now_date - right_month <= timedelta(0):
                left.find_element(By.CLASS_NAME, "el-icon-arrow-left").click()
            else:
                right.find_element(By.CLASS_NAME, "el-icon-arrow-right").click()
            WebDriverWait(right_month_ele, 10).until(lambda ele: ele.text != right_month_str)
        last_days = left.find_elements(By.XPATH, ".//td[contains(@class, 'available')]")
        last_days[0].click()
        last_days[-1].click()

    def __date_select_3(self, submodule: WebElement):
        """日期选择3:会员新增情况统计表的那个类型日期选择控件"""
        condition = EC.element_to_be_clickable((By.XPATH, ".//input[@placeholder='请选择日期']"))
        select = WebDriverWait(submodule, 10).until(condition)
        select.click()
        WebDriverWait(self._driver, 10).until(EC.visibility_of_element_located((By.CLASS_NAME, "ant-calendar")))
        last_month, calendar_ele = self.__locate_last_month()
        self._driver.find_element(By.XPATH, f".//td[@title='{last_month}1日']").click()
        WebDriverWait(self._driver, 10).until(EC.invisibility_of_element(calendar_ele))
        last_month, calendar_ele = self.__locate_last_month()
        pattern = (By.XPATH, ".//td[@class='ant-calendar-cell ant-calendar-last-day-of-month']")
        self._driver.find_element(*pattern).click()
        WebDriverWait(self._driver, 10).until(EC.invisibility_of_element(calendar_ele))

    def __locate_last_month(self):
        """日期选择3的日历控件定位到上个月"""
        now_date = datetime.today()
        while True:
            calendar_month_ele = self._driver.find_element(By.CLASS_NAME, "ant-calendar-ym-select")
            calendar_month_str = calendar_month_ele.text
            year_int, month_int = calendar_month_str.split("年")
            year_int = int(year_int)
            month_int = int(month_int[:-1])
            if year_int == now_date.year and month_int == now_date.month:
                self._driver.find_element(By.CLASS_NAME, "ant-calendar-prev-month-btn").click()
                WebDriverWait(calendar_month_ele, 10).until(lambda ele: ele.text != calendar_month_str)
                break
            calendar_month = datetime.strptime(calendar_month_str, '%Y年%m月')
            if now_date - calendar_month <= timedelta(0):
                self._driver.find_element(By.CLASS_NAME, "ant-calendar-prev-month-btn").click()
            else:
                self._driver.find_element(By.CLASS_NAME, "ant-calendar-next-month-btn").click()
            WebDriverWait(calendar_month_ele, 10).until(lambda ele: ele.text != calendar_month_str)
        return calendar_month_ele.text, calendar_month_ele

    def __search(self, submodule: WebElement):
        submodule.find_element(By.XPATH, "//button[contains(., '查询')]").click()

    def __wait_search_finnsh_1(self, submodule: WebElement):
        """等待查询内容显示1:类似综合收款统计的表格类型"""
        pattern = (By.XPATH, ".//div[@class='ant-spin-nested-loading']//div[@class='ant-spin-container']")
        WebDriverWait(submodule, 10).until(EC.presence_of_element_located(pattern))

    def __wait_search_finnsh_2(self):
        """等待查询内容显示2:类似储值消费汇总表的表格类型"""
        condition = EC.presence_of_element_located((By.CLASS_NAME, 'el-loading-parent--relative'))
        WebDriverWait(self._driver, 10).until_not(condition)

    def __clear_excel(self, name):
        """清理已存在的文件"""
        file_names = os.listdir(self._download_path)
        for file_name in file_names:
            if name not in file_name:
                continue
            path = os.path.join(self._download_path, file_name)
            os.remove(path)
            print(f"清理文件:{path}")

    def __download_direct(self, submodule: WebElement, name):
        """直接导出文件"""
        submodule.find_element(By.XPATH, ".//span[text()='导出']/parent::button").click()
        file_name = self.__wait_download(name)
        self._download_file[name] = file_name

    def __download_autotrophy_detail(self, submodule: WebElement, name):
        submodule.find_element(By.XPATH, ".//span[text()='导出']/parent::button").click()
        condition = EC.visibility_of_element_located((By.XPATH, "//div[@id='rcDialogTitle0']/../.."))
        dialog = WebDriverWait(self._driver, 10).until(condition)
        for content in ["菜品明细", "支付明细", "优惠明细"]:
            select_ele = dialog.find_element(By.XPATH, f".//span[text()='{content}']/preceding-sibling::span")
            if "ant-checkbox-checked" not in select_ele.get_attribute("class"):
                select_ele.click()
            WebDriverWait(select_ele, 2).until(lambda ele: "ant-checkbox-checked" in ele.get_attribute("class"))
        dialog.find_element(By.XPATH, ".//span[text()='确 定']/parent::button").click()
        file_name = self.__wait_download(name)
        self._download_file[name] = file_name

    def __wait_download(self, name):
        """等待下载"""
        st = time.time()
        while True:
            if (time.time() - st) > self._download_timeout:
                raise Exception("Waiting download timeout.")
            file_names = os.listdir(self._download_path)
            for file_name in file_names:
                if name in file_name:
                    break
            else:
                continue
            break
        if file_name.endswith(".xlsx"):
            return file_name
        print(f"文件已经在下载:{file_name}")
        file_name = file_name.replace(".crdownload", "")
        return file_name

    def __synthesize_income_condition(self, submodule: WebElement):
        """综合收款统计的查询条件"""
        element = submodule.find_element(By.XPATH, ".//span[text()='按 日']/..")
        if "isSelected" in element.get_attribute("class"):
            print("当前统计周期已经是：按日")
        else:
            element.click()
        element = submodule.find_element(By.XPATH, ".//span[text()='按业务小类统计']/..")
        if " ant-checkbox-wrapper-checked" in element.get_attribute("class"):
            print("已经是按业务小类统计")
        else:
            element.click()

    def __pay_settlement_condition(self, submodule: WebElement):
        """支付结算表的查询条件"""
        element = submodule.find_element(By.XPATH, ".//span[text()='交易日期']/..")
        if "isSelected" in element.get_attribute("class"):
            print("当前统计方式已经是：交易日期")
        else:
            element.click()

    def __toggle_old(self):
        """切换到旧版本:会员新增情况统计表的功能"""
        c1 = EC.visibility_of_element_located((By.XPATH, "//span[text()='切换回老版']"))
        c2 = EC.visibility_of_element_located((By.XPATH, "//span[text()='切换回新版']"))
        version = WebDriverWait(self._driver, 10).until(EC.any_of(c1, c2))
        if version.text == "切换回老版":
            version.click()
        else:
            print("当前已经在老版本")


class ElemeData:
    """饿了么相关留存表的处理"""

    def __init__(self) -> None:
        self.wb = load_workbook(GOL.save_path.eleme_bill)

    def save(self):
        self.wb.save(GOL.save_path.eleme_bill)

    def del_useless_sheet(self):
        """删除没用的分表"""
        for name in self.wb.sheetnames:
            if name in ["账单汇总", "外卖订单明细", "保险相关业务账单明细", "赔偿单"]:
                continue
            print(f"删除分表:{name}")
            sheet = self.wb[name]
            self.wb.remove(sheet)

    def billing_summary(self):
        """账单汇总分表的处理"""
        ws = self.wb["账单汇总"]
        header = [ws.cell(1, i).value for i in range(1, ws.max_column + 1)]
        assert header == ['结算入账ID', '门店ID', '门店名称', '账单日期', '结算金额', '结算日期', '账单类型']
        insert_data = []
        # 删除行
        i = 1
        while True:
            i += 1
            if i > ws.max_row:
                break
            bill_type = ws.cell(i, 7).value
            if bill_type == "外卖单":
                continue
            ws.delete_rows(i)
            i -= 1
        # 提取保险相关业务账单明细的数据
        insurance_data = {}
        insurance_ws = self.wb["保险相关业务账单明细"]
        header = [insurance_ws.cell(1, i).value for i in range(1, insurance_ws.max_column + 1)]
        date_i = header.index("账单日期")
        amount_i = header.index("结算金额")
        for row in insurance_ws.iter_rows(min_row=2, values_only=True):
            date_str = row[date_i]
            amount = row[amount_i]
            insurance_data[date_str] = amount
        insert_data.append(["保险金额", insurance_data])
        # 提取抖音渠道佣金明细的数据
        tiktok_ws = self.wb["抖音渠道佣金明细"]
        if tiktok_ws.max_row > 1:
            header = [ws.cell(1, i).value for i in range(1, tiktok_ws.max_column + 1)]
            date_i = header.index("账单日期")
            amount_i = header.index("结算金额")
            tiktok_data = {}
            for row in tiktok_ws.iter_rows(min_row=2, values_only=True):
                date_str = row[date_i]
                amount = row[amount_i]
                tiktok_data[date_str] = amount
            insert_data.append(["抖音渠道佣金", tiktok_data])
        # 插入列，并填入数据
        for j, datas in enumerate(insert_data):
            name, values = datas
            col_i = 6 + j
            ws.insert_cols(col_i, 1)
            ws.cell(1, col_i, name)
            for i in range(2, ws.max_row + 1):
                date_str = ws.cell(i, 4).value
                if date_str not in values:
                    continue
                ws.cell(i, col_i, values.pop(date_str))
            assert len(values) == 0
        # 计算总和
        col_i = 6 + len(insert_data)
        ws.insert_cols(col_i, 1)
        ws.cell(1, col_i, "结算金额合计")
        for i in range(2, ws.max_row + 1):
            ws.cell(i, col_i, f"=SUM(E{i}:{get_column_letter(col_i - 1)}{i})")
        # 最后一行插入合计行
        sum_row_i = ws.max_row + 1
        ws.cell(sum_row_i, 1, "合计")
        for i in range(5, 7 + len(insert_data)):
            range_start = f"{get_column_letter(i)}2"
            range_end = f"{get_column_letter(i)}{sum_row_i - 1}"
            ws.cell(sum_row_i, i, f"=SUM({range_start}:{range_end})")

    def take_out(self):
        """处理外卖订单明细分表"""
        ws = self.wb["外卖订单明细"]
        assert ws.cell(1, 12).value == "结算金额"
        assert ws.cell(1, 15).value == "菜价"
        assert ws.cell(1, 17).value == "技术服务费"
        assert ws.cell(1, 25).value == "智能满减津贴"
        assert ws.cell(1, 29).value == "商户配送费"
        assert ws.cell(1, 43).value == "商户活动补贴"
        assert ws.cell(1, 46).value == "商户配送费活动补贴"

        # 插入两列
        ws.insert_cols(13, 2)
        # 计算结算金额
        ws.cell(1, 13, "结算")
        for i in range(2, ws.max_row + 1):
            ws.cell(i, 13, f"=Q{i}+AE{i}+AS{i}+S{i}+AV{i}+AA{i}")
        # 计算差额
        ws.cell(1, 14, "差额")
        for i in range(2, ws.max_row + 1):
            ws.cell(i, 14, f"=M{i}-L{i}")
        # 复制差额不为0的数据到新表
        compensate_data = []
        for row in ws.iter_rows(min_row=2, values_only=True):
            balance = Decimal(str(row[16])) + Decimal(str(row[18])) + Decimal(str(row[26])) + Decimal(str(row[30])) \
                + Decimal(str(row[44])) + Decimal(str(row[47])) - Decimal(str(row[11]))
            if balance == Decimal(str(0)):
                continue
            compensate_data.append(row)
        if len(compensate_data) != 0:
            compensate_ws = self.wb.create_sheet("赔偿单")
            for i in range(1, ws.max_column + 1):
                compensate_ws.cell(1, i, ws.cell(1, i).value)
            for i, row in enumerate(compensate_data, start=2):
                for j, value in enumerate(row, start=1):
                    compensate_ws.cell(i, j, value)
                compensate_ws.cell(i, 13, f"=Q{i}+AE{i}+AS{i}+S{i}+AV{i}+AA{i}")
                compensate_ws.cell(i, 14, f"=M{i}-L{i}")
        # 计算合计
        sum_row_i = ws.max_row + 1
        ws.cell(sum_row_i, 1, "合计")
        for i in range(13, ws.max_column - 1):
            col_str = get_column_letter(i)
            range_start = f"{col_str}2"
            range_end = f"{col_str}{sum_row_i - 1}"
            ws.cell(sum_row_i, i, f"=SUM({range_start}:{range_end})")


class DaDaAutotrophy:
    """达达门店订单明细的处理"""

    def __init__(self) -> None:
        self.wb = load_workbook(GOL.save_path.autotrophy_dada)

    def save(self):
        self.wb.save(GOL.save_path.autotrophy_dada)

    def autotrophy_order(self):
        """自营订单外卖的数据处理"""
        # 获取需要保存的数据
        ws = self.wb["1"]
        header = [ws.cell(1, i).value for i in range(1, ws.max_column + 1)]
        order_source_i = header.index("订单来源编号")
        order_status_i = header.index("订单状态")
        save_data = [header]
        for row in ws.iter_rows(min_row=2, values_only=True):
            order_source = row[order_source_i]
            order_status = row[order_status_i]
            if "自营外卖" not in order_source:
                continue
            elif order_status != "已完成":
                continue
            save_data.append(list(row))
        # 将配送距离的值从文本改为数值
        distance_i = header.index("配送距离")
        for i in range(1, len(save_data)):
            save_data[i][distance_i] = float(save_data[i][distance_i])
        # 保存数据
        ws = self.wb.create_sheet("自营外卖订单（不含自配送）")
        for i, row in enumerate(save_data, start=1):
            for j, value in enumerate(row, start=1):
                ws.cell(i, j, value)
        # 插入数据
        ws.insert_cols(15, 1)
        ws.cell(1, 15, "配送区间应收客户运费")
        assert ws.cell(1, 14).value == "配送距离"
        for i in range(2, ws.max_row + 1):
            formula = f"=IF(N{i}<=2000,2,IF(N{i}<=3000,3,IF(N{i}<=4000,4,IF(N{i}<=5000,6,IF(N{i}<=6000,7,"
            formula += f"IF(N{i}<=7000,9,IF(N{i}<=9000,10,IF(N{i}<=10000,12,IF(N{i}<=12000,15,"
            formula += f"IF(N{i}<=14000,18,IF(N{i}<=15000,22,"")))))))))))"
            ws.cell(i, 15, formula)
        ws.insert_cols(40, 1)
        ws.cell(1, 40, "商户支付配送费")
        assert ws.cell(1, 39).value == "运费账户消耗"
        for i in range(2, ws.max_row + 1):
            ws.cell(i, 40, f"=AM{i}-O{i}")
        # 添加合计行
        sum_row_i = ws.max_row + 1
        ws.cell(sum_row_i, 1, "合计")
        for i in range(15, 41):
            range_start = f"{get_column_letter(i)}2"
            range_end = f"{get_column_letter(i)}{sum_row_i - 1}"
            ws.cell(sum_row_i, i, f"=SUM({range_start}:{range_end})")


class MeiTuanAutotrophy:
    """美团自营外卖表的处理"""

    def __init__(self) -> None:
        self.wb = load_workbook(GOL.save_path.autotrophy_meituan)

    def save(self):
        self.wb.save(GOL.save_path.autotrophy_meituan)

    def order_detail(self):
        """订单明细分表的处理"""
        # 获取需要另存为的数据
        ws = self.wb["订单明细"]
        header = [ws.cell(3, i).value for i in range(1, ws.max_column + 1)]
        order_source_i = header.index("订单来源")
        order_status_i = header.index("订单状态")
        distribution_mode_i = header.index("配送方式")
        save_data = [header]
        for row in ws.iter_rows(min_row=4, values_only=True):
            if row[order_source_i] != "自营外卖":
                continue
            elif row[order_status_i] != "已完成":
                continue
            elif row[distribution_mode_i] != "自配送(达达配送)":
                continue
            save_data.append(row)
        # 保存数据
        ws = self.wb.create_sheet("自营外卖订单明细（不含自配送）")
        for i, row in enumerate(save_data, start=1):
            for j, value in enumerate(row, start=1):
                ws.cell(i, j, value)
        # 添加合计行
        sum_row_i = ws.max_row + 1
        ws.cell(sum_row_i, 1, "合计")
        assert header[14:19] == ['订单金额（元）', '顾客应付（元）', '支付合计（元）', '订单优惠（元）', '订单收入（元）']
        for i in range(15, 20):
            col_str = get_column_letter(i)
            range_start = f"{col_str}2"
            range_end = f"{col_str}{sum_row_i - 1}"
            ws.cell(sum_row_i, i, f"=SUM({range_start}:{range_end})")


class TalkOutData:
    """外卖订单汇总表的数据处理"""

    def __init__(self) -> None:
        self.wb = load_workbook(GOL.save_path.take_out)

    def save(self):
        self.wb.save(GOL.save_path.take_out)

    def collect_eleme(self):
        """汇总饿了么的数据"""
        with xw.App(visible=False) as app:
            with app.books.open(GOL.save_path.eleme_bill) as wb:
                ws = wb.sheets["外卖订单明细"]
                last_cell = ws.used_range.last_cell
                max_row, max_col = last_cell.row, last_cell.column
                header: List = ws[f"A1:{get_column_letter(max_col)}1"].value
                should_income = ws[f"{get_column_letter(header.index('菜价') + 1)}{max_row}"].value
                service_cost = ws[f"{get_column_letter(header.index('技术服务费') + 1)}{max_row}"].value
                time_raise = ws[f"{get_column_letter(header.index('履约-时段加价') + 1)}{max_row}"].value
                distance_raise = ws[f"{get_column_letter(header.index('履约-距离加价') + 1)}{max_row}"].value
                price_raise = ws[f"{get_column_letter(header.index('履约-价格加价') + 1)}{max_row}"].value
                activty_subsidy = ws[f"{get_column_letter(header.index('商户活动补贴') + 1)}{max_row}"].value +\
                    ws[f"{get_column_letter(header.index('智能满减津贴') + 1)}{max_row}"].value
                chargeback = -ws[f"{get_column_letter(header.index('差额') + 1)}{max_row}"].value
                other = ws[f"{get_column_letter(header.index('商户配送费活动补贴') + 1)}{max_row}"].value

                ws = wb.sheets["账单汇总"]
                last_cell = ws.used_range.last_cell
                max_row, max_col = last_cell.row, last_cell.column
                header: List = ws[f"A1:{get_column_letter(max_col)}1"].value
                insurance = -ws[f"{get_column_letter(header.index('保险金额') + 1)}{max_row}"].value

        ws = self.wb[f"{GOL.last_month.year}年Bread饿了么"]
        row_i = [ws.cell(i, 1).value.strftime("%y.%m") for i in range(4, 16)]
        row_i = row_i.index(GOL.last_month.strftime("%y.%m")) + 4
        ws.cell(row_i, 2, should_income)  # 应收
        ws.cell(row_i, 3, service_cost)  # 抽佣服务费
        ws.cell(row_i, 4, time_raise)  # 时段加价
        ws.cell(row_i, 5, distance_raise)  # 距离加价
        ws.cell(row_i, 6, price_raise)  # 价格加价
        ws.cell(row_i, 7, activty_subsidy)  # 商户承担活动补贴
        ws.cell(row_i, 8, chargeback)  # 部分退单-申请退单金额
        ws.cell(row_i, 9, other)  # 其他(商家自行配送补贴)
        ws.cell(row_i, 11, insurance)  # 保险费

    def collect_autotrophy(self):
        """汇总自营外卖的数据"""
        with xw.App(visible=False) as app:
            with app.books.open(GOL.save_path.autotrophy_meituan) as wb:
                ws = wb.sheets["自营外卖订单明细（不含自配送）"]
                last_cell = ws.used_range.last_cell
                max_row, max_col = last_cell.row, last_cell.column
                header: List = ws[f"A1:{get_column_letter(max_col)}1"].value

                order_num = max_row - 2
                col_str = get_column_letter(header.index("结账方式") + 1)

                member_consume = len([data for data in ws[f"{col_str}1:{col_str}{max_row - 1}"].value if "会员卡" in data])
                order_amount = ws[f"{get_column_letter(header.index('订单金额（元）') + 1)}{max_row}"].value
                discount_amount = ws[f"{get_column_letter(header.index('订单优惠（元）') + 1)}{max_row}"].value
            with app.books.open(GOL.save_path.autotrophy_dada) as wb:
                ws = wb.sheets["自营外卖订单（不含自配送）"]
                last_cell = ws.used_range.last_cell
                max_row, max_col = last_cell.row, last_cell.column
                header: List = ws[f"A1:{get_column_letter(max_col)}1"].value

                client_freight = ws[f"{get_column_letter(header.index('配送区间应收客户运费') + 1)}{max_row}"].value
                business_freight = ws[f"{get_column_letter(header.index('商户支付配送费') + 1)}{max_row}"].value

        ws = self.wb[f"{GOL.last_month.year}年Bread自营外卖"]
        row_i = [ws.cell(i, 1).value.strftime("%Y.%m") for i in range(3, 15)]
        row_i = row_i.index(GOL.last_month.strftime("%Y.%m")) + 3
        ws.cell(row_i, 2, order_num)  # 订单数量
        ws.cell(row_i, 3, member_consume)  # 会员消费订单数
        ws.cell(row_i, 5, order_amount)  # 订单金额
        ws.cell(row_i, 6, discount_amount)  # 折扣金额(含会员、运费及商品）
        ws.cell(row_i, 7, client_freight)  # 客人支付运费
        ws.cell(row_i, 8, business_freight)  # 商户承担运费


def init_chrome(path, download_path, user_path):
    service = Service(path)
    options = Options()
    options.add_argument('--log-level=3')
    options.add_argument(f'user-data-dir={user_path}')
    options.add_experimental_option('prefs', {
        "download.default_directory": download_path,  # 指定下载目录
    })
    chrome_driver = Chrome(service=service, options=options)
    return chrome_driver


def crawler_main(chrome_path, download_path, user_path):
    """网站抓取的主方法"""
    try:
        print("启动浏览器，开始爬虫抓取")
        with init_chrome(chrome_path, download_path, user_path) as driver:
            meituan = MeiTuanCrawler(driver, download_path)
            print("登录并打开美团网站")
            meituan.login()
            print("下载综合营业统计表")
            meituan.download_synthesize_operate()
            print("下载自营外卖/自提订单明细表")
            meituan.download_autotrophy()
            print("下载综合收款统计表的数据")
            meituan.download_synthesize_income()
            print("下载支付结算表的相关数据")
            meituan.download_pay_settlement()
            print("下载储值消费汇总表的数据")
            meituan.download_store_consume()
            print("下载会员新增情况统计表的相关数据")
            meituan.download_member_addition()
            print("从美团网站爬虫导出EXCEL文件已完成")
            dada = DadaCrawler(driver, download_path)
            print("登入并打开达达网站")
            dada.login()
            print("下载门店明细报表")
            dada.download_store_report()
            print("从达达网站爬虫导出EXCEL文件已完成")
            print("等待网站上的EXCEL文件导出完毕,并移动EXCEL到相应位置")
            meituan.wait_download_finnish()
            dada.wait_download_finnish()
    except Exception:
        print(traceback.format_exc())
        return
    print("爬虫部分已完成")


def operation_detail_main(template_path):
    """营业明细表汇总的主方法"""
    print("定义数据保存类")
    writer = GetOperateDetail(template_path)
    print("读取综合营业统计表的数据")
    writer.read_general_business()
    print("读取综合收款统计表的数据")
    writer.read_general_collection()
    print("读取储值消费汇总表的数据")
    writer.read_store_consume()
    print("读取会员新增情况统计表的相关数据")
    writer.read_newly_increased()
    print("读取支付结算表的相关数据")
    writer.read_pay_settlement()
    print("将数据写入定义好的excel模板中")
    writer.write_and_save()
    print("营业明细表的数据已汇总完成")


def eleme_main():
    """饿了么数据的表格处理的主方法"""
    eleme_data = ElemeData()
    print("账单汇总分表的处理")
    eleme_data.billing_summary()
    print("处理外卖订单明细分表")
    eleme_data.take_out()
    print("删除不保留的分表")
    eleme_data.del_useless_sheet()
    print("保存文件")
    eleme_data.save()
    print("饿了么数据的留存表已整理完毕")


def dada_autotrophy_main():
    """达达门店订单明细处理的主方法"""
    excel = DaDaAutotrophy()
    print("处理自营外卖订单")
    excel.autotrophy_order()
    print("保存文件")
    excel.save()
    print("达达门店明细的留存表已整理完毕")


def meituan_autotrophy_main():
    """美团自营外卖表处理的主方法"""
    excel = MeiTuanAutotrophy()
    print("处理订单明细分表")
    excel.order_detail()
    print("保存文件")
    excel.save()
    print("美团自营外卖表的留存表已整理完毕")


def take_out_main():
    """外卖收入汇总表的主方法"""
    take_out = TalkOutData()
    print("汇总饿了么的数据")
    take_out.collect_eleme()
    print("汇总自营外卖的数据")
    take_out.collect_autotrophy()
    print("保存文件")
    take_out.save()
    print("外卖收入的数据已汇总完成")


def main():
    """主方法"""
    operate_detail_template = r"E:\NewFolder\xiayun\营业明细表模板.xlsx"
    user_path = r'C:\Users\Administrator\AppData\Local\Google\Chrome\User Data'
    chrome_path = r'E:\NewFolder\chromedriver_mac_arm64_114\chromedriver.exe'
    download_path = r"D:\Download"
    # 定义保存名称
    save_folder = os.path.dirname(operate_detail_template)
    date_str = GOL.last_month.strftime("%Y%m")
    GOL.save_path.operate_detail = os.path.join(save_folder, f"营业明细表{date_str}.xlsx",)
    GOL.save_path.synthesize_operate = os.path.join(save_folder, f"综合营业统计{date_str}.xlsx")
    GOL.save_path.synthesize_income = os.path.join(save_folder, f"综合收款统计{date_str}.xlsx")
    GOL.save_path.store_consume = os.path.join(save_folder, f"储值消费汇总表{date_str}.xlsx")
    GOL.save_path.member_addition = os.path.join(save_folder, f"会员新增情况统计表{date_str}.xlsx")
    GOL.save_path.pay_settlement = os.path.join(save_folder, f"支付结算{date_str}.xlsx")
    GOL.save_path.take_out = os.path.join(save_folder, f"外卖收入汇总表{date_str[:4]}.xlsx")
    GOL.save_path.eleme_bill = os.path.join(save_folder, f"账单明细{date_str}.xlsx")
    GOL.save_path.autotrophy_meituan = os.path.join(save_folder, f"自营外卖{date_str}(美团后台).xlsx")
    GOL.save_path.autotrophy_dada = os.path.join(save_folder, f"自营外卖{date_str}(达达).xlsx")
    # 从美团网站上下载相关EXCEL文件
    crawler_main(chrome_path, download_path, user_path)
    # 汇总营业明细表
    operation_detail_main(operate_detail_template)
    # 饿了么导出数据的处理
    eleme_main()
    # 美团自营外卖表的处理
    meituan_autotrophy_main()
    # 达达门店订单明细的处理
    dada_autotrophy_main()
    # 汇总外卖收入表
    take_out_main()


if __name__ == "__main__":
    main()
