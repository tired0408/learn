import socket
import re
import time


def CRC(message):
    """
    生成序号1-8数据的CRC校验
    :param message: 序号1-8数据的ascii码
    :return:返回CRC校验码(ascii码格式)
    """
    CRCFull = 0xFFFF
    for i in range(len(message)):
        CRCFull = CRCFull ^ message[i]
        for j in range(8):
            CRCLSB = CRCFull & 0x0001
            CRCFull = (CRCFull >> 1) & 0x7FFF
            if (CRCLSB == 1):
                CRCFull = CRCFull ^ 0xA001
    return CRCFull.to_bytes(2, byteorder="little")


host = '172.16.115.179'
port = 58258

net_init_00 = b"\xFF\xFF\xFF\xFF\xFF\xFF\x00\x00\x00\x00"
start_01 = b'\x78'  # 包头一帧数据的起始，1bytes
ver_02 = b'\x34'  # 协议版本,1bytes
addr_03 = b'\x01\x00'  # 控制卡地址,2bytes
cmd_04 = b'\x24'  # 通信命令,1bytes
ident_05 = b'\x00\x00\x00\x00'  # 识别标志,4bytes
frame_06 = b'\x00\x00\x00\x00'  # 帧计数,4bytes
len_07 = b"\x02\x00"  # data字段的数据长度,2bytes
data_08 = b"\x00\x0f"  # 数据,nbytes
check_09 = b'\x35\x16'  # CRC 校验,序号1-8数据的CRC 效验,2bytes
end_10 = b'\xA5'  # 一帧结束标志,1bytes


def set_screen_brightness(value):
    """
    设置屏幕亮度
    :param value: int类型，范围1-15
    :return: 返回TCP通信需要的ascii码
    """
    cmd_04 = b"\x24"
    ident_05 = b'\x00\x00\x00\x00'
    data_08 = b"\x00" + bytes([value])
    len_07 = b"\x02\x00"
    check_09 = CRC(start_01 + ver_02 + addr_03 + cmd_04 + ident_05 + frame_06 + len_07 + data_08)
    rd = net_init_00 + start_01 + ver_02 + addr_03 + cmd_04 + ident_05 + frame_06 + len_07 + data_08 + check_09 + end_10
    return rd


def change_distance(id_value, content_value):
    """
    修改车辆的行驶距离
    :param value: int类型，范围0-999
    :return: 返回TCP通信需要的ascii码
    """
    cmd_04 = b"\x29"
    ident_05 = b'\xbc\xfd\x00\x00'
    id = int(id_value).to_bytes(2, byteorder="little")  # 字符分区ID,2byte
    code = b"\x01"  # 编码方式0:unicode,1:gb2312, 1byte
    show_type = b"\x02"  # 显示方式,0:保存数据模式，2：立即显示，1byte
    index = b"\x06"  # 字符串索引，1byte
    color = b"\x01"  # 颜色,1byte
    content = content_value.encode("gb2312")
    # content = b""  # 字符串, nbyte
    # for i in content_value:
    #     content += b"\x00"
    #     content += i.encode("ascii")
    length = len(content).to_bytes(2, byteorder="little")  # 长度，2byte
    data_08 = id + code + show_type + index + color + length + content
    len_07 = len(data_08).to_bytes(2, byteorder="little")
    check_09 = CRC(start_01 + ver_02 + addr_03 + cmd_04 + ident_05 + frame_06 + len_07 + data_08)
    return net_init_00 + start_01 + ver_02 + addr_03 + cmd_04 + ident_05 + frame_06 + len_07 + data_08 + check_09 + end_10


tcpCliSock = socket.socket(socket.AF_INET, socket.SOCK_STREAM)  # 开启套接字
tcpCliSock.connect((host, port))  # 连接到服务器
start = time.time()
send_data = change_distance(36448, "我爱中航软件")
# send_data = set_screen_brightness(15)
print("send_data:", re.sub(r"(?<=\w)(?=(?:\w\w)+$)", " ", send_data.hex()))
tcpCliSock.send(send_data)  # 发送信息.
response = tcpCliSock.recv(1024)  # 接受返回信息
print("response:", re.sub(r"(?<=\w)(?=(?:\w\w)+$)", " ", response.hex()))
print("total time:%d" % (time.time() - start))
tcpCliSock.close()
