# 港口码头VO类
class Terminal(object):
    terminal_base_info = {"HTD": ["海天码头2号闸口", 118.091838, 24.511304],  # 使用的1号闸口
                          "HTM": ["海通码头", 118.045131, 24.449933],
                          "HRD": ["海润码头", 118.016415, 24.456572],
                          "YHT": ["厦门远海集装箱码头有限公司", 117.977207, 24.465757],
                          "SYT": ["厦门嵩屿集装箱码头有限公司", 118.03867, 24.45302],
                          "H.C": ["厦门国际货柜码头有限公司", 118.027067, 24.455853],
                          "XHD": ["新海达码头", 117.975202, 24.466041]}

    def __init__(self, code):
        """
        :param code: 港口代码
        """
        self.code = code
        self.name = self.terminal_base_info[code][0]
        self.bd_lng = self.terminal_base_info[code][1]
        self.bd_lat = self.terminal_base_info[code][2]
