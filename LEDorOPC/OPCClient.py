#coding=utf-8
import socket
import struct
from OPCDataConstant import *
import json
"""
opc连接类，基于python3
"""
class OpcModel(object):
    """
    TCP协议发送规范：
    帧起始，1byte, 0x21
    协议版本，1byte, 0x00
    请求类型：1byte，0x00获取数据，0x01控制PLC状态
    请求数据列表，2byte, 01表示有没有，序号根据OPCDataConstant,例如有SLING_HEGHT，LOCK_STATUS，则为1001 0000 0000 0000
    帧结束，1byte,0x57
    TCP协议接收规范：
    帧起始，1byte, 0x09
    协议版本，1byte, 0x00
    结果标识：1byte, 0x00错误, 0x01成功
    数据长度：2byte
    数据：nbyte
    """
    tcp_start = b"\x21"  # 帧起始
    tcp_version = b"\x00"  # 协议版本
    tcp_end = b"\x57"  # 帧结束
    def __init__(self, host="127.0.0.1", port=51112):
        self.host = host
        self.port = port
        self.client = socket.socket(socket.AF_INET, socket.SOCK_STREAM) # 开启套接字
        self.client.connect((host, port)) # 连接到服务器
        # self.client.settimeout(0.3)

    def getData(self, name_list):
        """
        根据列表，获取返回数据
        :param name_list:
        :return:
        """
        # 根据规范构造发送的字节码，并获取返回的字节码
        request_type = b"\x00"
        data_int = 0x0000
        for name in name_list:
            data_int = data_int | (1 << name)
        data_byte = struct.pack('<H', data_int)
        send_info = self.tcp_start+self.tcp_version+request_type+data_byte+self.tcp_end
        self.client.send(send_info)
        response = self.client.recv(1024)
        data_list = []
        # 解析返回的字节码
        if response[2] == 0x00:
            error_len = struct.unpack("<H", response[3:5])[0]
            error_byte = response[5:5+error_len]
            print("error:%s" % error_byte.decode("gb2312"))
        elif response[2] == 0x01:
            return_len = struct.unpack("<H", response[3:5])[0]
            return_str = response[5:5+return_len].decode("gb2312")
            return_json = json.loads(return_str)
            return_list = return_json["data"]
            for name in name_list:
                data_list.append(return_list[name])
        return data_list

    def is_continue_fall(self, value):
        """
        控制PLC的状态（是否允许吊具继续下降）
        :param value:
        :return:
        """
        request_type = b"\x01"
        data_byte = struct.pack("?", value)
        send_info = self.tcp_start + self.tcp_version + request_type + data_byte + self.tcp_end
        self.client.send(send_info)
        response = self.client.recv(1024)
        print("Success,Response:%s" % response)


    def __del__(self):
        self.client.close()


if __name__ == "__main__":
    opc_model = OpcModel("127.0.0.1", 51112)
    while True:
        print(opc_model.getData([SLING_HEGHT,
                                    PACKING_TYPE,
                                    LOAD_OR_UNLOAD,
                                    LOAD_INDEX,
                                    FORWARD_DIRECTION,
                                    IS_SINGLE_BOX,
                                    LOCK_STATUS,
                                    DRIVE_STATUS,
                                    CONTROL_ON,
                                    SMALL_CAR_LOCATION]))