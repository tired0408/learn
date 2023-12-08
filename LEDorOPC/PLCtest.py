import socket
import re
import time


host = '172.16.115.179'
port = 58258

def build_send_data():
    # 218Header的帧内容
    command_type = b"\x11"  # 11命令，19响应，12报文
    identification_number = b"\x00"   # 发送报文的序列号，从00H-FFH依次反复
    # 根据所用安川 PLC 的 CPU号，如果不是安川 PLC 系列，则 channel NO.统一设置为 00H。
    destination_channel_no = b"\x00"
    source_channel_no = b"\x11"
    data_length = b"\x13\x00"  # 218Header 和 Application data 的总长字节数（低位在前，高位在后）
    # Application data 的帧内容
    length = b"\x07\x00"  # 命令的长度
    mfc = b"\x20"  # MFC总是20H
    function_codes = b"\x01"  # 功能代码，01H读线圈状态，05H更改单线圈状态，06H写单寄存器，09H读保持寄存器，0BH写保持寄存器
    cpu_no = b"\x12"  # 前4位是目的CPU号，后4位是源CPU号
    # reference_no = b"\x00"
    reference_no = b"\x00\x00"  # 设定要操作的 PLC 中线圈的地址，占用两个字节
    # number_of_coils = b"\x01"
    number_of_coils = b"\x01\x00"  # 读取的线圈的数量，占用两个字节
    data = command_type+identification_number+destination_channel_no+source_channel_no+data_length+length+mfc+ \
                function_codes+cpu_no+reference_no+number_of_coils
    return data

# tcpCliSock = socket.socket(socket.AF_INET, socket.SOCK_STREAM)  # 开启套接字
# tcpCliSock.connect((host, port))  # 连接到服务器
# start = time.time()
# send_data = build_send_data()
# print("send_data:", re.sub(r"(?<=\w)(?=(?:\w\w)+$)", " ", send_data.hex()))
# tcpCliSock.send(send_data)  # 发送信息
# response = tcpCliSock.recv(1024)  # 接受返回信息
# print("response:", re.sub(r"(?<=\w)(?=(?:\w\w)+$)", " ", response.hex()))
# print("total time:%d" % (time.time() - start))
# tcpCliSock.close()
print(build_send_data().hex())


