"""
拦截浏览器响应数据的方法，并识别验证码图片，启动命令：
mitmdump -s mitmproxy.py
"""
import base64
import requests
from mitmproxy.http import HTTPFlow


class BaiduOCRApi:

    def __init__(self) -> None:
        self.url = self.get_baidu_api()

    def get_baidu_api(self):
        api_key = "lQgPjrptNrsaBxVcEgZIPScq"
        secret_key = "BGZDwlifkIp0EF4WTgENUy8hluHL4gZg"

        url = "https://aip.baidubce.com/oauth/2.0/token?grant_type=client_credentials&"
        url += f"client_id={api_key}&client_secret={secret_key}"
        headers = {
            'Content-Type': 'application/json',
            'Accept': 'application/json'
        }
        response = requests.request("POST", url, headers=headers, data="")
        data: dict = response.json()
        access_token = data.get("access_token")
        request_url = "https://aip.baidubce.com/rest/2.0/ocr/v1/numbers"
        request_url = request_url + "?access_token=" + access_token
        return request_url

    def recongnize(self, value):
        response = requests.post(self.url, data={"image": value})
        if response:
            data = response.json()
            number = "".join([str(value["words"]) for value in data["words_result"]])
            return number


def response(flow: HTTPFlow):

    if "code.jsp" in flow.request.url or "captcha.jpg" in flow.request.url:
        print("-" * 200)
        print(flow.request.url)
        img = base64.b64encode(flow.response.content)
        result = ocr_api.recongnize(img)
        print(result)
        rd = requests.get(r"http://localhost:8557/ocr", params={"value": result}, timeout=0.5)
        print(rd)
        print("-" * 200)


ocr_api = BaiduOCRApi()
