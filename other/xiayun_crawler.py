"""
爬虫抓取夏云所需的销售数据报表,并填入已有的数据表中
"""
import os
import time
import shutil
import warnings
import calendar
import datetime
import traceback
import pandas as pd
from selenium.webdriver import Chrome
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.remote.webelement import WebElement
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import NoSuchElementException, TimeoutException
from typing import Dict, Tuple, List
from datetime import timedelta
from openpyxl import load_workbook


class PerDayData:
    """每日数据表"""

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


class DataToExcel:

    def __init__(self, path, data: Dict[str, PerDayData], days, save_name) -> None:
        self.data = data
        self.days = days
        self.save_name = save_name
        self.save_folder = os.path.dirname(path)
        self.wb = load_workbook(path)
        self.ws = self.wb.active

    def write_and_save(self):
        # TODO 调整EXCEL，新增或删除行，暂时人工处理
        # 设置开始、结束时间
        self.ws.cell(2, 2, self.days[0])
        self.ws.cell(3, 2, self.days[-1])
        # 设置每天的数据
        for index, day in enumerate(self.days):
            row_index = 5 + index
            day_data = self.data[day]
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
        self.wb.save(os.path.join(self.save_folder, self.save_name["结果"]))

    def read_general_business(self):
        """读取综合营业统计表的相关数据"""
        path = os.path.join(self.save_folder, self.save_name["综合营业统计"])
        data = pd.read_excel(path, header=None)
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
            day_data = self.data[day_str]
            day_data.cash = row[cach_i]
            day_data.wechat = row[wechat_i]
            day_data.eat_in = sum([row[i] for i in eat_in_i_list])
            day_data.ele_me = row[ele_me_i]
            day_data.ele_me_free = row[ele_me_free_i]
            day_data.other_free = row[other_free_i]
            day_data.pubilc_relation_income = 0 if pubilc_relation_income_i is None else row[pubilc_relation_income_i]

    def read_general_collection(self):
        """读取综合收款统计表的相关数据"""
        path = os.path.join(self.save_folder, self.save_name["综合收款统计"])
        data = pd.read_excel(path, header=None)
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
            day_data = self.data[day_str]
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
        for _, day_data in self.data.items():
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
        path = os.path.join(self.save_folder, self.save_name["储值消费汇总表"])
        data = pd.read_excel(path, header=None)
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
            day_data = self.data[day_str]
            day_data.main_consume = row[5]
            day_data.gift_consume = row[6]
            day_data.main_paid = row[2]
            day_data.gift_paid = row[3]

    def read_newly_increased(self):
        """"读取会员新增情况统计表的相关数据"""
        path = os.path.join(self.save_folder, self.save_name["会员新增情况统计表"])
        data = pd.read_excel(path, header=None)
        assert data.iloc[2, 0] == "日期", "表格发生变化，请联系管理员"
        assert data.iloc[2, 1] == "合计", "表格发生变化，请联系管理员"
        for row in data.iloc[3:].itertuples():
            if row[1] == "合计":
                break
            day_str: str = row[1][2:]
            day_str = day_str.replace("-", ".")
            day_data = self.data[day_str]
            day_data.new_member = row[2]

    def read_pay_settlement(self):
        """读取支付结算表的相关数据"""
        path = os.path.join(self.save_folder, self.save_name["支付结算"])
        data = pd.read_excel(path, header=None)
        assert data.iloc[2, 2] == "结算日期", "表格发生变化，请联系管理员"
        assert data.iloc[2, 3] == "交易金额(元)", "表格发生变化，请联系管理员"
        assert data.iloc[2, 5] == "手续费(元)", "表格发生变化，请联系管理员"
        for row in data.iloc[3:len(self.days) + 3].itertuples():
            day_str = row[3]
            if pd.isna(day_str):
                break
            day_str: str = day_str[2:]
            day_str = day_str.replace("-", ".")
            day_data = self.data[day_str]
            day_data.hand_charge = row[6]
            day_data.settlement_amount = row[4]


class Crawler:

    def __init__(self, exe_path, download_path, user_path, save_folder) -> None:
        self.download_timeout = 60
        self.download_path = download_path
        self.download_file = {}
        self.save_folder = save_folder
        self.driver = self.init_chrome(exe_path, download_path, user_path)
        self.action = ActionChains(self.driver)

    def __del__(self):
        self.driver.quit()

    @staticmethod
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
        self.driver.get(r"https://pos.meituan.com")

    def general_collection_condition(self, container: WebElement):
        """综合收款统计表的导出条件"""
        element = container.find_element(By.XPATH, ".//span[text()='按 日']/..")
        if "isSelected" in element.get_attribute("class"):
            print("当前统计周期已经是：按日")
        else:
            element.click()
        element = container.find_element(By.XPATH, ".//span[text()='按业务小类统计']/..")
        if " ant-checkbox-wrapper-checked" in element.get_attribute("class"):
            print("已经是按业务小类统计")
        else:
            element.click()

    def pay_settlement_condition(self, container: WebElement):
        """下载支付结算表的相关数据"""
        element = container.find_element(By.XPATH, ".//span[text()='交易日期']/..")
        if "isSelected" in element.get_attribute("class"):
            print("当前统计方式已经是：交易日期")
        else:
            element.click()

    def store_consume_condition(self):
        """下载储值消费汇总表的数据"""
        # 设置查询日期
        self.driver.find_element(By.CLASS_NAME, "el-range-input").click()
        pattern = (By.XPATH, ".//div[@class='el-picker-panel el-date-range-picker el-popper']")
        WebDriverWait(self.driver, 10).until(EC.visibility_of_element_located(pattern))
        left, right = self.driver.find_element(*pattern).find_elements(By.XPATH, "./*/*/*")
        now_date = datetime.datetime.today()
        while True:
            try:
                right.find_element(By.CLASS_NAME, "today")
                break
            except NoSuchElementException:
                pass
            right_month_ele = right.find_element(By.XPATH, "./div/div")
            right_month_str = right_month_ele.text
            right_month = datetime.datetime.strptime(right_month_str, '%Y 年 %m 月')
            if now_date - right_month <= timedelta(0):
                left.find_element(By.CLASS_NAME, "el-icon-arrow-left").click()
            else:
                right.find_element(By.CLASS_NAME, "el-icon-arrow-right").click()
            WebDriverWait(right_month_ele, 10).until(lambda ele: ele.text != right_month_str)
        last_days = left.find_elements(By.XPATH, ".//td[contains(@class, 'available')]")
        last_days[0].click()
        last_days[-1].click()
        # 查询
        self.driver.find_element(By.XPATH, ".//span[contains(text(), '查询')]/parent::button").click()
        WebDriverWait(self.driver, 10).until_not(EC.presence_of_element_located(
            (By.CLASS_NAME, 'el-loading-parent--relative')))

    def newly_increased_method(self):
        """下载会员新增情况统计表的相关数据"""
        c1 = EC.visibility_of_element_located((By.XPATH, "//span[text()='切换回老版']"))
        c2 = EC.visibility_of_element_located((By.XPATH, "//span[text()='切换回新版']"))
        version = WebDriverWait(self.driver, 10).until(EC.any_of(c1, c2))
        if version.text == "切换回老版":
            version.click()
        else:
            print("当前已经在老版本")
        pattern = (By.XPATH, "//input[@placeholder='请选择日期']")
        WebDriverWait(self.driver, 10).until(EC.visibility_of_element_located(pattern))
        self.driver.find_element(*pattern).click()
        WebDriverWait(self.driver, 10).until(EC.visibility_of_element_located((By.CLASS_NAME, "ant-calendar")))
        last_month, calendar_ele = self.__locate_last_month()
        self.driver.find_element(By.XPATH, f"//td[@title='{last_month}1日']").click()
        WebDriverWait(self.driver, 10).until(EC.invisibility_of_element(calendar_ele))
        last_month, calendar_ele = self.__locate_last_month()
        pattern = (By.XPATH, "//td[@class='ant-calendar-cell ant-calendar-last-day-of-month']")
        self.driver.find_element(*pattern).click()
        WebDriverWait(self.driver, 10).until(EC.invisibility_of_element(calendar_ele))
        # 查询
        self.driver.find_element(By.XPATH, "//button[contains(., '查询')]").click()
        pattern = (By.XPATH, "//form//div[@class='ant-tablex ant-tablex_hasFooter']//div[@class='ant-spin-container']")
        WebDriverWait(self.driver, 10).until(EC.presence_of_element_located(pattern))

    def market_center_data(self, menu_name, name, condition_method):
        """下载营销中心的子模块的相关EXCEL文件"""
        # 进入营销中心
        module = self.__enter_main_module("营销中心")
        # 进入子模块
        if self.__get_active_submodule_name(module) != name:
            self.__hover_and_click(module, menu_name, name)
        name2url = {
            "储值消费汇总表": '/web/crm-smart/report/summary-store',
            "会员新增情况统计表": '/web/member/statistic/member-increase#/'
        }
        pattern = (By.XPATH, f".//iframe[@data-current-url='{name2url[name]}']")
        WebDriverWait(module, 60).until(EC.visibility_of_element_located(pattern))
        iframe = module.find_element(*pattern)
        self.driver.switch_to.frame(iframe)
        # 各个模块的条件设置
        condition_method()
        # 导出
        self.__download_excel(name, self.driver)
        self.driver.switch_to.default_content()

    def report_center_data(self, menu_name, name, condition_method=None):
        """下载报表中心的子模块的相关EXCEL文件"""
        # 进入报表中心
        module = self.__enter_main_module("报表中心")
        # 进入子模块
        submodule = self.__get_statement_submodule(module)
        # 如果不是目标子模块，则进入子模块
        if self.__get_active_submodule_name(module) != name:
            self.__hover_and_click(module, menu_name, name)
            # 等待原有模块消失，即新模块显示
            WebDriverWait(submodule, 10).until(lambda ele: ele.get_attribute("style") == "display: none;")
            submodule = self.__get_statement_submodule(module)
        # 美团的日期选择会自动跳回去，怀疑是根据日历中标签aria-selected="true"来恢复的。所以设置input的value没有用
        # start, end = driver.find_elements(By.XPATH, "//div[@data-key='timeRange']//input[@placeholder='请选择日期']")
        # driver.execute_script("arguments[0].value = arguments[1];", start, "2024/03/01")
        # driver.execute_script("arguments[0].setAttribute('value', arguments[1]);", start, "2024/03/01")
        # driver.execute_script("arguments[0].value = arguments[1];", end, "2024/03/31")
        # driver.execute_script("arguments[0].setAttribute('value', arguments[1]);", end, "2024/03/31")
        # 日期选择上个月
        pattern = (By.XPATH, ".//input[@placeholder='请选择日期']")
        WebDriverWait(submodule, 10).until(EC.visibility_of_element_located(pattern))
        submodule.find_element(*pattern).click()
        pattern = (By.XPATH, "//div[@class='ant-calendar-footer']//span[text()='上月']")
        WebDriverWait(self.driver, 10).until(EC.visibility_of_element_located(pattern))
        self.driver.find_element(*pattern).click()
        WebDriverWait(self.driver, 10).until_not(EC.visibility_of_element_located(pattern))
        # 特定条件的设置
        if condition_method is not None:
            condition_method(submodule)
        # 查询数据并导出
        submodule.find_element(By.XPATH, ".//span[text()='查询']/parent::button").click()
        pattern = (By.XPATH, ".//div[@class='auto2-page-slot_body']//div[@class='ant-spin-container']")
        WebDriverWait(submodule, 10).until(EC.presence_of_element_located(pattern))
        self.__download_excel(name, submodule)

    def wait_download_finnish(self, save_name):
        # 等到下载完成
        for _, file_name in self.download_file.items():
            st = time.time()
            while True:
                if (time.time() - st) > self.download_timeout:
                    raise Exception("Waiting download timeout.")
                if os.path.exists(os.path.join(self.download_path, file_name)):
                    break
        # 移动下载文件
        for key, file_name in self.download_file.items():
            src = os.path.join(self.download_path, file_name)
            dst = os.path.join(self.save_folder, save_name[key])
            shutil.move(src, dst)
            self.download_file[key] = dst

    def __download_excel(self, name, submodule: WebElement):
        """导出excel文件并移动到相应文件夹"""
        # 清理文件
        file_names = os.listdir(self.download_path)
        for file_name in file_names:
            if name not in file_name:
                continue
            path = os.path.join(self.download_path, file_name)
            os.remove(path)
            print(f"清理文件:{path}")
        # 导出文件
        submodule.find_element(By.XPATH, ".//span[text()='导出']/parent::button").click()
        file_name = self.__wait_download(name)
        self.download_file[name] = file_name

    def __enter_main_module(self, name):
        """进入主模块的方法：运营中心、营销中心、库存管理、报表中心"""
        # 获取当前容器显示的模块内容
        pattern = (By.XPATH, "//div[@class='main-app']/div[@style='display: block;'][./*]")
        WebDriverWait(self.driver, 10).until(EC.presence_of_all_elements_located(pattern))
        current_module = self.driver.find_element(*pattern)
        # 进入所定模块
        pattern = (By.XPATH, f"//header//span[text()='{name}']/..")
        WebDriverWait(self.driver, 10).until(EC.visibility_of_element_located(pattern))
        element = self.driver.find_element(*pattern)
        if "active-first-menu" in element.get_attribute("class"):
            print(f"已在{name}模块")
            return current_module
        element.click()
        # 等待原有模块消失，即新模块显示
        WebDriverWait(current_module, 10).until(lambda ele: ele.get_attribute("style") == "display: none;")
        # 获取新模块
        pattern = (By.XPATH, "//div[@class='main-app']/div[@style='display: block;'][./*]")
        current_module = self.driver.find_element(*pattern)
        try:
            current_module.find_element(By.XPATH, f".//div[@role='tablist']//span[text()='{name}首页']")
        except NoSuchElementException:
            raise Exception("进入模块出现问题")
        return current_module

    def __get_statement_submodule(self, module: WebElement):
        """获取报表中心激活的子模块"""
        pattern = (By.XPATH, ".//div[@id='__root_wrapper_rms-report']//div[@style='display: block;']")
        WebDriverWait(module, 10).until(EC.presence_of_all_elements_located(pattern))
        current_submodule = module.find_element(*pattern)
        return current_submodule

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
                self.action.move_to_element(menu).perform()
                pattern = (By.XPATH, f"//ul[@id='{menu_id}$Menu']//li[text()='{name}']")
                WebDriverWait(self.driver, 1).until(EC.visibility_of_element_located(pattern))
                self.driver.find_element(*pattern).click()
            except TimeoutException:
                continue
            break

    def __locate_last_month(self):
        """日历控件定位上个月"""
        now_date = datetime.datetime.today()
        while True:
            calendar_month_ele = self.driver.find_element(By.CLASS_NAME, "ant-calendar-ym-select")
            calendar_month_str = calendar_month_ele.text
            year_int, month_int = calendar_month_str.split("年")
            year_int = int(year_int)
            month_int = int(month_int[:-1])
            if year_int == now_date.year and month_int == now_date.month:
                self.driver.find_element(By.CLASS_NAME, "ant-calendar-prev-month-btn").click()
                WebDriverWait(calendar_month_ele, 10).until(lambda ele: ele.text != calendar_month_str)
                break
            calendar_month = datetime.datetime.strptime(calendar_month_str, '%Y年%m月')
            if now_date - calendar_month <= timedelta(0):
                self.driver.find_element(By.CLASS_NAME, "ant-calendar-prev-month-btn").click()
            else:
                self.driver.find_element(By.CLASS_NAME, "ant-calendar-next-month-btn").click()
            WebDriverWait(calendar_month_ele, 10).until(lambda ele: ele.text != calendar_month_str)
        return calendar_month_ele.text, calendar_month_ele

    def __wait_download(self, name):
        """等待下载"""
        st = time.time()
        day_str = datetime.date.today().strftime("%Y-%m-%d")
        name = f"{name}_{day_str}"
        while True:
            if (time.time() - st) > self.download_timeout:
                raise Exception("Waiting download timeout.")
            file_names = os.listdir(self.download_path)
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


def init_data() -> Tuple[Dict[str, PerDayData], List[str]]:
    """初始化存储数据"""
    # 获取上个月所有日期
    today = datetime.date.today()
    first_day_of_this_month = today.replace(day=1)
    last_day_of_last_month = first_day_of_this_month - timedelta(days=1)
    year, month = last_day_of_last_month.year, last_day_of_last_month.month
    num_days = calendar.monthrange(year, month)[1]
    data, days = {}, []
    for day in range(1, num_days + 1):
        day_str = datetime.date(year, month, day).strftime("%y.%m.%d")
        data[day_str] = PerDayData()
        days.append(day_str)
    # 定义存储名字
    date_str = last_day_of_last_month.strftime("%Y%m")
    save_name = {
        "结果": f"营业明细表{date_str}.xlsx",
        "综合营业统计": f"综合营业统计{date_str}.xlsx",
        "综合收款统计": f"综合收款统计{date_str}.xlsx",
        "储值消费汇总表": f"储值消费汇总表{date_str}.xlsx",
        "会员新增情况统计表": f"会员新增情况统计表{date_str}.xlsx",
        "支付结算": f"支付结算{date_str}.xlsx"
    }
    print(f"已定义了:{date_str}的数据")
    return data, days, save_name


def main(excel_path, chrome_path, download_path, user_path):
    save_data, last_days, save_name = init_data()
    try:
        print("定义所需服务")
        crawler = Crawler(chrome_path, download_path, user_path, os.path.dirname(excel_path))
        print("登录并打开网站")
        crawler.login()
        print("下载综合营业统计表")
        crawler.report_center_data("营业报表", "综合营业统计")
        print("下载综合收款统计表的数据")
        crawler.report_center_data("收款报表", "综合收款统计", crawler.general_collection_condition)
        print("下载支付结算表的相关数据")
        crawler.report_center_data("收款报表", "支付结算", crawler.pay_settlement_condition)
        print("下载储值消费汇总表的数据")
        crawler.market_center_data("数据报表", "储值消费汇总表", crawler.store_consume_condition)
        print("下载会员新增情况统计表的相关数据")
        crawler.market_center_data("用户", "会员新增情况统计表", crawler.newly_increased_method)
        print("等待网站上的EXCEL文件导出完毕,并移动EXCEL到相应位置")
        crawler.wait_download_finnish(save_name)
        print("退出浏览器")
        crawler.driver.quit()
    except Exception:
        print(traceback.format_exc())
        return
    print("定义数据保存类")
    warnings.simplefilter("ignore")  # 忽略pandas使用openpyxl读取excel文件的警告
    writer = DataToExcel(excel_path, save_data, last_days, save_name)
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
    print("程序运行完成")


if __name__ == "__main__":
    set_file_path = r"E:\NewFolder\xiayun\template.xlsx"
    set_user_path = r'C:\Users\Administrator\AppData\Local\Google\Chrome\User Data'
    set_chrome_path = r'E:\NewFolder\chromedriver_mac_arm64_114\chromedriver.exe'
    set_download_path = r"D:\Download"
    date_str = datetime.datetime.today()
    main(set_file_path, set_chrome_path, set_download_path, set_user_path)
