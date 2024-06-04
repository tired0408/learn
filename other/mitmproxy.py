"""
拦截浏览器响应数据的方法，并识别验证码图片，启动命令：
mitmdump -s mitmproxy.py
"""
import time
import base64
import socket
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
            if "words_result" not in data:
                return ""
            number = "".join([str(value["words"]) for value in data["words_result"]])
            return number


class CaptchaSocketClient:
    """验证码服务的socket端"""

    def __init__(self) -> None:
        self.sock = socket.socket(socket.AF_INET, socket.SOCK_STREAM)

    def reconnect(self):
        """重新连接"""
        for _ in range(5):
            try:
                print("Attempting to reconnect...")
                self.sock = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
                self.sock.connect(('localhost', 12345))
                print("Reconnected successfully.")
                break
            except socket.error as e:
                print("Reconnection failed:", e)
                print("Retrying in 5 seconds...")
                time.sleep(5)
        else:
            raise Exception("重新连接服务端失败多次")

    def send(self, value: str):
        """发送信息"""
        for _ in range(10):
            try:
                self.sock.sendall(value.encode())
                break
            except ConnectionResetError as e:
                print(f"Connection was reset:{e}")
                self.reconnect()
                continue
            except socket.error as e:
                print(f"Socket is not connected: {e}")
                self.sock.connect(('localhost', 12345))
                continue
        else:
            raise Exception("尝试发送多次失败")


def response(flow: HTTPFlow):

    if "code.jsp" in flow.request.url or "captcha.jpg" in flow.request.url:
        print("-" * 200)
        print(flow.request.url)
        img = base64.b64encode(flow.response.content)
        result = ocr_api.recongnize(img)
        print(f"识别结果:{result}")
        if len(result) < 4:
            result += "00000"
            print(f"识别结果加上00000:{result}")
        sock.send(result)
        print("-" * 200)


ocr_api = BaiduOCRApi()
sock = CaptchaSocketClient()
