import socket
import time
"""
LED屏UDP连接类
"""
class ConnectLEDwithUDP(object):

    def __init__(self, host, port):
        self.host = host
        self.port = port
        self.tcpCliSock = socket.socket(socket.AF_INET, socket.SOCK_DGRAM)
        self.tcpCliSock.settimeout(2)
        self.msg_index = None  # 初始化节目
        self.distance = None  # 初始化距离
        self.lane = None  # 初始化车道
        self.text = None  # 初始化预留文本框的信息

    def changeScreenShow(self, direction=None, distance=0, lane=1, sling_state=None, text=None, opc_connect=True):
        """
        修改LED屏显示
        :param opc_connect: opc连接状态，默认True(连接)
        :param text: 文字内容
        :param sling_state: 吊具状态，默认False(危险)
        :param direction: 方向指令，"right"向右，"left"向左，"in_place"到位，"not_work"未作业
        :param distance: 距离数值，范围0-999
        :param lane: 车道信息，范围0-99
        :return:
        """
        msg_index = 0
        is_change_msg = False
        # 切换节目
        if direction is not None:
            if direction == "left":
                msg_index = 1
            elif direction == "right":
                msg_index = 2
            elif direction == "in_place":
                msg_index = 3
            elif direction == "not_work":
                msg_index = 7
        if sling_state is not None:
            if not sling_state:
                msg_index = 4
            else:
                msg_index = 5
        if not opc_connect:
            msg_index = 6
        if text is not None:
            msg_index = 8

        if self.msg_index != msg_index:
            msg = "##001000%d@" % msg_index
            self.tcpCliSock.sendto(msg.encode(), (self.host, self.port))
            self.tcpCliSock.recvfrom(1024)
            self.msg_index = msg_index
            is_change_msg = True
        # 修改距离，%C颜色，1红色2绿色3黄色
        if 0<self.msg_index<4:
            if self.distance != distance or is_change_msg:
                msg =  "!#001%ZI01%ZH9999%C2%AH2%AV2{:0>3d}$$".format(distance)
                self.tcpCliSock.sendto(msg.encode(), (self.host, self.port))
                self.tcpCliSock.recvfrom(1024)
                self.distance = distance
        # 修改车道
        elif 3<self.msg_index<6 or is_change_msg:
            if self.lane != lane:
                msg =  "!#001%ZI01%ZH9999%C3%AH2%AV2{}$$".format(lane)
                self.tcpCliSock.sendto(msg.encode(), (self.host, self.port))
                self.tcpCliSock.recvfrom(1024)
                self.lane = lane
        # 修改预留文本框信息
        elif self.msg_index == 8:
            if self.text != text or is_change_msg:
                msg =  "!#001%ZI01%ZH9999%F16%C2%AH2%AV2{}$$".format(text)
                self.tcpCliSock.sendto(msg.encode("gb2312"), (self.host, self.port))
                self.tcpCliSock.recvfrom(1024)
                self.text = text
    def __del__(self):
        self.tcpCliSock.close()


if __name__ == "__main__":
    host1 = '192.168.100.197'
    port1 = 5005
    con_led = ConnectLEDwithUDP(host1, port1)
    start = time.time()
    con_led.changeScreenShow(direction="right")
    print("total time:%d" % (time.time() - start))



