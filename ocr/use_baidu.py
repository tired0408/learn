import requests
import base64
import os
import shutil
import re
import traceback
def contain_zh(word):
    """
    判断字符串中是否包含中文
    :return: {bool} 包含返回True， 不包含返回False
    """
    for ch in word:
        if u'\u4e00' <= ch <= u'\u9fff':
            return True
    return False

class BaiduAPI(object):

    def __init__(self, access_key, secrect_key):
        self.access_key = access_key
        self.secrect_key = secrect_key


    def get_access_token(self):
        # client_id 为官网获取的AK， client_secret 为官网获取的SK
        host = "https://aip.baidubce.com/oauth/2.0/token?grant_type=client_credentials&client_id=%s&client_secret=%s" \
               % (self.access_key, self.secrect_key)
        response = requests.get(host)
        if response:
            result_data = response.json()
            return result_data["access_token"]

    def ocr_request(self, image_path, is_accurate=False):
        """
        文字识别
        :param is_accurate: 是否为高精度版
        :param image_path:
        :return:
        """
        data = {}
        if not is_accurate:
            url = "https://aip.baidubce.com/rest/2.0/ocr/v1/general_basic"
            data["language_type"] = "ENG"
        else:
            url = "https://aip.baidubce.com/rest/2.0/ocr/v1/accurate_basic"
        access_token = self.get_access_token()
        data["access_token"] = access_token
        with open(image_path, 'rb') as f:
            image_data = f.read()
            base64_data = base64.b64encode(image_data)
        data["image"] = base64_data

        response = requests.post(url, data)
        if response:
            result_data = response.json()
            return result_data["words_result"]


def use_baidu_api(raw_path):
    # TODO 需替换掉自己百度云,文字识别应用的API key和secret key
    ak = "Es8EHH22fXVkwbMIhlLKem4h"
    sk = "SEISpzEj1IODSGzneQ6uNogMlIG6iAbi"
    baidu_api = BaiduAPI(ak, sk)
    # 定义存储文件夹结构
    save_path = raw_path+"_deal"
    if not os.path.exists(save_path):
        os.mkdir(save_path)
    can_identify_path = os.path.join(save_path, "identify")
    if not os.path.exists(can_identify_path):
        os.mkdir(can_identify_path)
    identify_null_path = os.path.join(save_path, "identify_null")
    if not os.path.exists(identify_null_path):
        os.mkdir(identify_null_path)
    identify_other_path = os.path.join(save_path, "identify_other")
    if not os.path.exists(identify_other_path):
        os.mkdir(identify_other_path)
    identify_double_path = os.path.join(save_path, "identify_double")
    if not os.path.exists(identify_double_path):
        os.mkdir(identify_double_path)
    # 开始识别图片
    filenames = os.listdir(raw_path)
    # 删掉某个文件之前的东西

    identify_other_string = ""
    can_identify_string = ""
    error_name = ""
    for name in filenames:
        ac_rs = None
        try:
            ac_rs = baidu_api.ocr_request(os.path.join(raw_path, name))
        except:
            traceback.format_exc()
            error_name += name+"\n"
        if ac_rs is None:
            error_name += name + "\n"
            continue
        if len(ac_rs) == 0:
            print("identify is null:", name)
            shutil.copy(os.path.join(raw_path, name), os.path.join(identify_null_path, name))
        elif len(ac_rs) > 1:
            print("identify double data:", name)
            shutil.copy(os.path.join(raw_path, name), os.path.join(identify_double_path, name))
        else:
            label = ac_rs[0]["words"]
            label = label.replace(" ","").replace("/","|").replace("\\","|")
            if re.search('[^a-zA-Z0-9|]', label) is not None:
                print("identify other:", name, label)
                shutil.copy(os.path.join(raw_path, name), os.path.join(identify_other_path, name))
                identify_other_string += "%s %s\n" % (name, label)
            else:
                print("can identify:", name, label)
                shutil.copy(os.path.join(raw_path, name), os.path.join(can_identify_path, name))
                can_identify_string += "%s %s\n" % (name, label)
    with open(os.path.join(identify_other_path, "label.txt"), "w+", encoding="gbk") as f:
        f.write(identify_other_string)
    with open(os.path.join(can_identify_path, "label.txt"), "w+", encoding="gbk") as f:
        f.write(can_identify_string)


if __name__ == "__main__":
    path = r"C:\Users\Administrator\Desktop\label"
    file_list = os.listdir(path)
    for fl in file_list:
        if fl.endswith("_deal"):
            continue
        if fl in ["-0-", "-1-", "-2-", "-3-", "-4-"]:
            continue
        use_baidu_api(os.path.join(path, fl))