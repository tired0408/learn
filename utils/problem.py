import numpy as np
import sys
import time
import os
import psutil
import cv2
import queue
"""
1. pipe_related: 频繁调用管道影响速度
2. volatile_variables: python的可变变量注意事项
3. del_about_memory_usage: 删除元素时候的内存占用情况
4. speed_about_numpy: 加上transpose(2, 0, 1)运行速度变快
5. inherit_writing: 继承的错误写法导致报错
"""
def pipe_related():
    # 以下代码会导致频繁调用管道，导致其他使用管道的地方，效率降低。原理未知。
    # 明显延迟出现在subprocess的pipe管道上，其他地方较不明显
    import subprocess as sp
    ffmpeg_pipe = sp.Popen("ps -ef", stdout=sp.PIPE, preexec_fn=os.setsid)
    test_queue = queue.Queue()
    while 1:
        if test_queue.empty():
            continue
        test_queue.get()
# 默认参数值在函数定义时只计算一次，这意味着修改参数的默认值将影响函数的所有后续调用。
def volatile_variables():
    def cache(lt=np.zeros((3, 3))):
        print(lt)
        lt[1,1] = 1
    cache()
    cache()
# 消减元素并不会释放内存
def del_about_memory_usage():
    # 全局变量内存
    bgr_image = np.zeros((1920, 1080, 3), dtype=np.uint8)  # linux减少,windows减少
    # bgr_image = cv2.imread("test.jpg")  # linux一段时间后减少,windows减少
    history_video = [bgr_image.copy() for i in range(250)]
    while 1:
        time.sleep(0.04)
        if len(history_video) > 0:
            history_video.pop(0)
        else:
            break
    # 类中的变量内存
    class A:
        def __init__(self):
            print("init")
            # windows减少，linux不减少
            # linux不减少可能为，linux判断后续还要用到内存，将该释放的内存转移过去
            self.bgr_image = cv2.imread("test.jpg")
            # self.bgr_image = np.zeros((1920,1080,3), dtype=np.uint8) # windows、linux减少,
            self.history_video = [self.bgr_image.copy() for i in range(250)]

        def del_list(self):
            print("del list start")
            while 1:
                time.sleep(0.04)
                if len(self.history_video) > 0:
                    self.history_video.pop(0)
                else:
                    break
            print("finnish del list")
    a = A()
    a.del_list()
    time.sleep(20)

def speed_about_numpy():
    # 加上transpose(2, 0, 1)运行速度变快，为什么？？？
    img = cv2.imread("test.jpg")
    np_img = img[:, :, ::-1]
    # np_img = np_img.transpose(2, 0, 1)
    np_img = np.ascontiguousarray(np_img)

def inherit_writing():
    class CameraBase:
        def __init__(self, size, fps=None):
            self.w, self.h = size
            self.fps = fps
    class CameraRead(CameraBase):
        def __init__(self, size, fps=None):
            super().__init__(size, fps=fps)
    class CameraSave(CameraBase):
        def __init__(self, size, fps=None):
            super().__init__(size, fps=fps)
    class CameraLive(CameraBase):
        def __init__(self, size):
            super().__init__(size)
    class CameraUse(CameraRead, CameraSave, CameraLive):
        def __init__(self, size, fps=None):
            fps = 15 if fps is None else fps
            CameraRead.__init__(self, size, fps=fps)
            CameraSave.__init__(self, size, fps=fps)
            CameraLive.__init__(self, size)
    camera = CameraUse([1920, 1080], fps=25)

