#coding=utf-8
"""
与OPC建立连接
"""
import OpenOPC

address_list = ["Applications.Application_1.gv_MainHoist.PosActualPosition",
                "Applications.Application_1.gv_Trolley.PosActualPosition",
                "Applications.Application_1.gv_Spreader.TelescopeSize20Feet",
                "Applications.Application_1.gv_Spreader.TelescopeSize40Feet",
                "Applications.Application_1.gv_Spreader.TelescopeSize45Feet",
                "Applications.Application_1.gv_Spreader.TelescopeTwin20",
                "Applications.Application_1.gv_Spreader.TwistlockUnLocked",
                "Applications.Application_1.gv_Spreader.TwistlockLocked",
                "Applications.Application_1.gv_Gantry.brakerelease",
                "Applications.Application_1.gv_CraneControl.ControlOn",
                "Applications.Application_1.gv_accs.SpareB9",
                "Applications.Application_1.gv_accs.SpareB10",
                "Applications.Application_1.gv_accs.SpareB11",
                "Applications.Application_1.MainHoist.v_RequestUp",
                "Applications.Application_1.MainHoist.v_RequestDown"]
opc_data = {}

# local test
# opc = OpenOPC.client()
# opc.connect(u"Matrikon.OPC.Simulation.1")
# def update_opc_data():
#     """
#     从OPC获取并更新数据
#     :return:
#     """
#     received_info = opc.read("Random.Int4")
#     for address in address_list:
#         opc_data[address] = received_info[0]
# local test
opc = OpenOPC.client()
opc.connect(u"ABB.AC800MC_OpcDaServer.3")
def update_opc_data():
    """
    从OPC获取并更新数据
    :return:
    """
    received_info = opc.read(address_list)
    for address, value, _, _ in received_info:
        opc_data[address] = value
"""
将获取的数据，通过TCP传送出去，基于python2.7
TCP协议发送规范：
帧起始，1byte, 0x09
协议版本，1byte, 0x00
结果标识：1byte, 0x00错误, 0x01成功
数据长度：2byte
数据：nbyte
帧结束，1byte,0x57
TCP协议接收规范：
帧起始，1byte, 0x21
协议版本，1byte, 0x00
请求类型：1byte，0x00获取数据，0x01控制PLC状态
请求数据列表，2byte, 01表示有没有，序号根据OPCDataConstant,例如有SLING_HEGHT，LOCK_STATUS，则为1001 0000 0000 0000
帧结束，1byte,0x57
"""
from OPCDataConstant import *
import socket
import struct
import json
import traceback
import os
import time
#  接受字符串的规范
rec_start = b"\x21"  # 帧起始
rec_version = b"\x00"  # 协议版本
rec_end = b"\x57"  # 帧结束
# 返回字符串的规范
ret_start = b"\x09"
ret_version = b"\x00"
ret_end = b"\x57"
def generate_error_byte(message):
    """
    根据错误信息，构造返回的字节串
    :param message: 错误信息
    :return:
    """
    error_msg = message.encode("gb2312")
    error_msg_len = struct.pack("<H", len(error_msg))
    return_byte = ret_start + ret_version + b"\x00" + error_msg_len + error_msg + ret_end
    return return_byte


def get_data_by_name(name):
    """
    根据名称，返回数据
    :return:
    """
    return_data = None
    if name == SLING_HEGHT:  # 吊具高度
        return_data = opc_data["Applications.Application_1.gv_MainHoist.PosActualPosition"]
    elif name == SMALL_CAR_LOCATION:
        return_data = opc_data["Applications.Application_1.gv_Trolley.PosActualPosition"]
    elif name == PACKING_TYPE:  # 装箱类型
        TelescopeSize20Feet = opc_data["Applications.Application_1.gv_Spreader.TelescopeSize20Feet"]
        TelescopeSize40Feet = opc_data["Applications.Application_1.gv_Spreader.TelescopeSize40Feet"]
        TelescopeSize45Feet = opc_data["Applications.Application_1.gv_Spreader.TelescopeSize45Feet"]
        if TelescopeSize20Feet:
            return_data = 20
        elif TelescopeSize40Feet:
            return_data = 40
        elif TelescopeSize45Feet:
            return_data = 45
    elif name == IS_SINGLE_BOX:  # 是否单箱
        TelescopeTwin20 = opc_data["Applications.Application_1.gv_Spreader.TelescopeTwin20"]
        return_data = not TelescopeTwin20
    elif name == LOCK_STATUS:   # 吊具开闭锁
        TwistlockUnLocked = opc_data["Applications.Application_1.gv_Spreader.TwistlockUnLocked"]
        TwistlockLocked = opc_data["Applications.Application_1.gv_Spreader.TwistlockLocked"]
        if TwistlockUnLocked:
            return_data = True
        elif TwistlockLocked:
            return_data = False
    elif name == DRIVE_STATUS:  # 桥吊行驶状态
        return_data = opc_data["Applications.Application_1.gv_Gantry.brakerelease"]
    elif name == CONTROL_ON:  # 控制合状态
        return_data = opc_data["Applications.Application_1.gv_CraneControl.ControlOn"]
    elif name == LOAD_OR_UNLOAD:  # 装卸船模式
        return_data = opc_data["Applications.Application_1.gv_accs.SpareB9"]
    elif name == LOAD_INDEX: # CPS集卡引导车道号
        return_data = opc_data["Applications.Application_1.gv_accs.SpareB10"]
    elif name == FORWARD_DIRECTION:  # CPS集卡行进方向选择
        return_data = opc_data["Applications.Application_1.gv_accs.SpareB11"]
    elif name == SLING_UP:  # 吊具状态
        v_RequestUp = opc_data["Applications.Application_1.MainHoist.v_RequestUp"]
        v_RequestDown = opc_data["Applications.Application_1.MainHoist.v_RequestDown"]
        if v_RequestUp:
            return_data = True
        elif v_RequestDown:
            return_data = False
    return return_data



def handle_client(client):
    """
    客户处理线程
    :param client:
    :return:
    """
    while True:
        request = client.recv(1024)  # 获取到的字节串
        print("[*] Received: %s" % request)
        if request:
            if request[0] != rec_start:
                client.send(generate_error_byte("帧起始不规范"))
            elif request[1] != rec_version:
                client.send(generate_error_byte("协议版本错误"))
            elif request[-1] != rec_end:
                client.send(generate_error_byte("帧结束不规范"))
            elif request[2] == b"\x00":  # 根据列表，返回对应数据
                update_opc_data()
                value_list = [None] * 16
                data_int = struct.unpack("<H", request[3:5])[0]
                for i in range(16):
                    if (data_int >> i) & 1 == 1:
                        value_list[i] = get_data_by_name(i)
                value_dict = {"data":value_list}
                value_str = json.dumps(value_dict)
                value_byte = value_str.encode("gb2312")
                value_len = struct.pack("<H", len(value_byte))
                return_value = ret_end+ret_version+b"\x01"+value_len+value_byte+ret_end
                client.send(return_value)  # 向客户端返回数据
            elif request[2] == b"\x01":  # 根据请求控制PLC状态
                # 操作PLC
                is_continue_fall = struct.unpack("?", request[3])[0]
                # opc.write(("Applications.Application_1.gv_accs.SpareB12", is_continue_fall))
                # opc.write(("Random.Boolean", is_continue_fall))
                # 向客户端返回结果
                value_byte = "OK".encode("gb2312")
                value_len = struct.pack("<H", len(value_byte))
                return_value = ret_end + ret_version + b"\x01" + value_len + value_byte + ret_end
                client.send(return_value)
        else:
            break
    client.close()

def save_error(value):
    """
    将程序异常导出到日志中
    :return:
    """
    rootDir = os.path.split(os.path.realpath(__file__))[0]
    blogpath = os.path.join(rootDir, 'errorlog.txt')
    timestamp = int(time.time())
    time_array = time.localtime(timestamp)  # 时间戳转时间数组（time.struct_time）
    time_str = time.strftime("%Y-%m-%d %H:%M:%S", time_array)  # 时间戳转字符串
    with open(blogpath, 'a+') as f:
        f.writelines("error: " + time_str + "\n" + value + "\n")

def main():
    try:
        bind_ip = "0.0.0.0"  # 监听所有可用的接口
        bind_port = 51112  # 非特权端口号都可以使用
        # AF_INET：使用标准的IPv4地址或主机名，SOCK_STREAM：说明这是一个TCP服务器
        server = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
        server.bind((bind_ip, bind_port))  # 服务器监听的ip和端口号
        print("[*] Listening on %s:%d" % (bind_ip, bind_port))
        server.listen(10)  # 让套接字进入被动监听状态
        #等待客户连接，连接成功后，将socket对象保存到client，将细节数据等保存到addr
        while True:
            try:
                client, addr = server.accept()
                print("[*] Acception connection from %s:%d" % (addr[0],addr[1]))
                handle_client(client)
                client.close()
            except Exception as err:
                err_str = "捕获监听错误" + "\n" + traceback.format_exc()
                print err_str
                save_error(err_str)
    except:
        err_str = "最外层捕获错误" + "\n" + traceback.format_exc()
        print err_str
        save_error(err_str)

if __name__ == "__main__":
    main()