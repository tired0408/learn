"""
爱股圈的数据抓取
"""
import os
import datetime
from crawler_util import init_chrome
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By
from selenium.webdriver.remote.webelement import WebElement

path = r"E:\NewFolder\aiguquan"
chrome_download = r"D:\Download"
chrome_path = os.path.join(path, r"..\chromedriver_mac_arm64_114\chrome114\App\Chrome-bin\chrome.exe")
chromedriver_path = os.path.join(path, r"..\chromedriver_mac_arm64_114\chromedriver.exe")
driver = init_chrome(chromedriver_path, chrome_download, chrome_path=chrome_path, is_proxy=False)
print("进入网站")
driver.get(r"https://www.aiguquan.com/vueapp/uc/chatting?forum_id=1465")
ele = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.XPATH, "//button[text()='只看圈主']/ancestor::div")))
if "bindThis" not in ele.get_attribute("class"):
    ele.click()
now_date = datetime.datetime.now()
for i in range(1, 7):
    driver.find_element(By.XPATH, "//div[contains(@class, 'el-date-editor')/input]").click()
print("关闭浏览器")
driver.quit()