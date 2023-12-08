#coding=utf-8
"""
OPC数据获取的索引，以及返回数据的说明，遇到OPC错误则统一返回None
"""
SLING_HEGHT = 0  # 吊具高度，与轨道为基准，float类型（单位：米）,4byte
PACKING_TYPE = 2  # 装箱类型，int类型，20、40、45尺, 2byte
LOAD_OR_UNLOAD = PACKING_TYPE+1 # 装卸船模式，int类型，0：NUll,1：装船，2：卸船
LOAD_INDEX = LOAD_OR_UNLOAD+1  # CPS集卡引导车道号，int类型
FORWARD_DIRECTION = LOAD_INDEX+1  # CPS集卡行进方向选择,1:集卡左进,2:集卡右进
IS_SINGLE_BOX = 6  # 是否单箱，bool类型，True单箱，False双箱, 1byte
LOCK_STATUS = IS_SINGLE_BOX + 1  # 吊具开闭锁，bool类型，True开锁，False闭锁, 1byte
DRIVE_STATUS = LOCK_STATUS + 1  # 桥吊行驶状态，bool类型，True行驶，False停车, 1byte
CONTROL_ON = DRIVE_STATUS + 1  # 控制合状态， bool类型，True开启，False关闭, 1byte
SMALL_CAR_LOCATION = SLING_HEGHT+1  # 小车位置，以海陆侧交接点为基准，往路侧为负，float类型（单位：米）, 4byte
SLING_UP = LOAD_INDEX + 1  # 吊具状态，bool类型。True：上升，False：下降



float_start_index = 0
int_start_index = 2
bool_start_index = 6

