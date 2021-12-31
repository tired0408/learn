#coding=utf-8
import random
import time
import datetime
import traceback
import pandas as pd
import matplotlib as mpl
import matplotlib.pyplot as plt
import os
import sys
import shutil
import torch
import cv2
import struct
import tensorflow as tf
def normal():
    """
    时间处理、字节处理、排序、指定GPU使用等。
    :return:
    """
    # 刷新当前行
    sys.stdout.write("需要刷新的内容")
    sys.stdout.flush()

    # try:
    #     raise Exception("自定义错误异常")
    # except Exception as e:
    #     print(traceback.format_exc())
    timestamp = int(time.time())
    time_array = time.localtime(timestamp)  # 时间戳转时间数组（time.struct_time）
    time_str = time.strftime("%Y-%m-%d %H:%M:%S", time_array)  # 时间戳转字符串
    time_array = time.strptime(time_str, "%Y-%m-%d %H:%M:%S")  # 字符串转时间数组（time.struct_time）
    timestamp = time.mktime(time_array)  # 字符串转化为时间戳
    print(timestamp)
    now_time = datetime.datetime.now()  # datetime.datetime 类型
    timestamp = int(now_time.timestamp())  # datetime.datetime类型时间转化为时间戳
    str = ["%02d:%02d" % (int(i/2), (i % 2)*30) for i in range(24)]  # 字符串格式化
    e_dc_indate = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")  # 存入mysql数据库的格式
    # 排序
    homework_value = {"a": ["wer",5], "b": ["sdf",2], "c": ["vwd",7], "d": ["btr",1]}
    sort_list = sorted(homework_value.items(), key=lambda x: x[1][1], reverse=True)
    # 获取文件夹下所有文件名
    os.listdir("../")
    # 编解码
    # 0x2211） 使用两个字节储存，高位字节是0x22，低位字节是0x11,人类读取为大端字节序，\x22\x11
    # 计算机电路先处理靠前的（较低的存储器地址），效率比较高，所以机器存储方式为，小端字节序，\x11\x22
    '12abc'.encode('ascii') # 字符串编码为字节码，b'12abc'
    bytes([1, 2, ord('1'), ord('2')]) # 数字或字符数组: b'\x01\x0212'
    bytes().fromhex('010210')  # 16进制字符串: b'\x01\x02\x10'
    bytes(map(ord, '\x01\x02\x31\x32')) # 16进制字符串: b'\x01\x0212'
    bytes([0x01, 0x02, 0x31, 0x32]) # 16进制数组:b'\x01\x0212'
    int.from_bytes(b"\x78", byteorder="big")  # 字节转换为整数: 120
    int(15).to_bytes(2, byteorder="big") # 整形转字节: b"\x00\x0f",大端(高位在前)
    # 二进制与十进制的转换
    bin(10)
    int('0b1010', 2)
    # python2，整形转字节,<小端, https://blog.csdn.net/qq_30638831/article/details/80421019
    c = 0x02000110
    b = struct.pack('<I', c)
    z = struct.unpack("<bHb", b)
    a = struct.pack("?", True)  # bool类型的
    # 求两个数组的交集，并集，差集
    set(a).intersection(set(b)) # 求交集
    set(a).difference(set(b)) # 求差集，a中有，b中没有的
    dict.get("abc", 0)  # 如果字典中无"abc"键，则返回0
    #######################
    ####### GPU相关
    #######################
    os.environ["CUDA_VISIBLE_DEVICES"] = "2" # 指定使用哪个GPU
    import GPUtil  # pip install gputil
    GPUs = GPUtil.getGPUs()
    print(GPUs[0].load)  # 0卡的使用率
    print(GPUs[0].memoryUsed)  # 使用的内存
    from onceorauto.misc.nvml import NVMLContext
    with NVMLContext() as ctx:
        for idx, dev in enumerate(ctx.devices()):
            print(dev.utilization()["gpu"])

def use_matlab():
    """
    使用matplotlib画图
    :return:
    """
    def draw_line():
        """
        画折线图
        :return:
        """
        mpl.rcParams['font.sans-serif'] = ['SimHei']  # 指定默认字体 SimHei为黑体
        mpl.rcParams['axes.unicode_minus'] = False  # 用来正常显示负号
        # plt.figure(figsize=(15, 15))  # 创建绘图对象
        plt.subplot(221)   # 构建2*2的方格，占据第一个位置
        x_value = range(24)
        y_value = [random.randint(0,10) for i in range(24)]
        plt.plot(x_value, y_value, marker='o')  # 在当前绘图对象进行绘图（两个参数是x,y轴的数据）,marker显示点
        # for a, b in zip(x_value, y_value):  # 增加点坐标数据
        #     plt.text(a, b, (a, b), ha='center', va='bottom', fontsize=10)
        plt.title("图像一")  # 标题
        plt.xlabel("hours")  # 横坐标标签
        plt.ylabel("total_num")  # 纵坐标标签
        # plt.savefig("time_feature.png")  # 保存图片
        plt.show()  # 显示图表
    def draw_rectangle():
        """
        画图像
        :return:
        """
        fig = plt.figure()  # 创建图
        ax = fig.add_subplot(111)  # 创建子图
        # ax = plt.gca() # 获得当前整张图表的坐标对象
        ax.invert_yaxis()  # y轴反向
        ax.xaxis.set_ticks_position('top')  # 将x轴的位置设置在顶部
        # 画矩形框，靠近原点的点坐标，长，宽
        rect = plt.Rectangle((0.1, 0.1), 0.5, 0.3, fill=False)
        ax.add_patch(rect)
        # 填充不规则图形
        # x = [1, 2, 2, 1]
        # y = [3, 3, 4, 4]open
        # rect = plt.fill(x, y, facecolor='g', alpha=0.5)
        # 画点
        # 这里，r表示red，+表示数据点的标记（marker)
        # -.表示线型，linewidth指定线宽，markersize表示标记的大小
        # plt.plot(x1, y1, 'r+-.', linewidth=1, markersize=12)
        plt.show()
    draw_rectangle()

def pandas_method():
    """
    pandas的使用方法，导出csv、excel文件
    :return:
    """
    columns = ['size', 'age', 'height']  # 文件标题
    data = [[0, 1, 1], [2, 1, 2]]  # 文件内容
    # 读取txt文件内容，返回DataFrame格式
    # 输入[文件地址；header数据的起始行；指定分隔符：指定哪一列为行索引;是否保留应该转化的缺省值列表]
    pd.read_csv(os.path.join("./", "label.txt"), header=0, sep=" ", index_col=0, keep_default_na=False)
    df = pd.DataFrame(data, index=range(len(data)), columns=columns)
    # 创建DataFrame
    test_dict = {'id': [1, 2, 3, 4, 5, 6], 'name': ['Alice', 'Bob', 'Cindy', 'Eric', 'Helen', 'Grace '],
                 'math': [90, 89, 99, 78, 97, 93], 'english': [89, 94, 80, 94, 94, 90]}
    df = pd.DataFrame(test_dict)
    # 输出csv文件[地址；行索引；列索引；编码格式；分隔符]
    df.to_csv("output.csv", header=0, index=False, encoding="utf-8", sep="\t")
    # loc基于行标签和列标签
    df.loc["a","name"]
    # iloc基于行索引和列索引
    df.iloc[0,1]
    # 根据条件对dataframe的内容进行修改
    max_min_scaler = lambda x: (x - np.min(x)) / (np.max(x) - np.min(x))
    df[3] = df[3].apply(max_min_scaler)
    # 根据索引删除行或列
    df.drop(index=[0,1]) # 删除行
    df.drop(columns=[1, 2]) # 删除列
    df.insert(0, "dir_path", np.array([0]*len(df)))  # 插入列[索引，列名，列数据]
    df = pd.concat([df, df], ignore_index=True)  # 合并dateframe[pd列表，是否忽视索引]
    # axis=0删除行，1删除列；去除有空值的行
    df = df[df[1].isna()]  # 获取第二列有空值的行
    df = df.dropna(axis=0, how="any")  # 删除任何有缺失值的行
    # 删除缺失数据, [维度0行1列，特定的列中含有的缺失值，是否修改源文件，all全为nan的行any带有nan的行]
    df.dropna(axis=0, subset=["Age", "Sex"], inplace=True, how="all")
    # 返回Series, 获取每行是否重复的结果。[subset选择比较的列;keep考虑重复值]
    # keep: "first":将第一个值视为唯一值,"last":将最后一个值视为唯一值, False将所有值视为重复值
    df.duplicated(subset=None, keep=False)
    # 去除重复的, keep表示是选择最前一项还是最后一项保留, inplace是直接在原来数据上修改还是保留一个副本，默认为False
    df.drop_duplicates(subset=None, keep='first', inplace=True)
    # 按比例随机抽取若干行,frac为比例
    df.sample(frac=1)
    # 根据条件筛选数据
    df = df[~df[3].isin([-1])]
    df = df[~df[5].str.contains(",", regex=True)]  #  令regex=True，则为正则表达式

import numpy as np
def numpy_method():
    """
    numpy的使用方法
    :return:
    """
    img = np.zeros((1920, 1080, 3), dtype=np.uint8)
    np.array([[10, 5], [20, 30], [70, 20], [50, 10]], np.int32)  # 数组转为numpy类型
    x = np.random.randn((5,3,5))  # 多维度数组
    # 从给定的一维数组中产生一个随机样本
    # a：1维数组。如果是整数，随机样本来自np.arange(a)
    # size：输出个数，如果是元组(m, n, k)，则m*n*k个样本生成。
    # replace：输出的数能否重复，False表示不能重复
    # p：输入a中每个数的输出概率。
    np.random.choice(a=5, size=2, replace=False, p=None)
    # 将多维数组降位一维
    # numpy.flatten()返回一份拷贝(copy)
    # numpy.ravel()返回的是视图(view)，影响原始矩阵
    x.ravel()
    x.flatten()
    # 取出最大值所对应的索引（索引值默认从0开始）,axis按第n维度进行计算
    np.argmax(x, axis=0)
    # 求范数,axis处理类型，keepdims是否保持矩阵的二维特性
    # ord(默认2)：2，表示二范数，即元素平方的和，再开根号；1一范数，元素绝对值的和；np.inf无穷范数，元素绝对值的最大值
    np.linalg.norm(x, axis=None, keepdims=False, ord=None)
    # 打乱数组， [数组]
    np.random.shuffle(x)
    # 根据条件返回索引值
    np.argwhere(x == 1)
    # 保存数组，读取数组
    np.save("numpy.npy", x)
    np.load("numpy.npy")
    # 将第1个维度和第二个维度交换
    x = x.swapaxes(0, 1)
    # 增加数组维度，原数组（5，3，5），增加后（5，3，1，5）
    y = x[:,:,np.newaxis,:]
    np.expand_dims(x, axis=2)
    # 数组经过行上的切片操作，会改变数组的连续性。关于连续性的内容
    # 第三个和第四个冒号——取遍图像的所有通道数，-1是反向取值,
    x = x[:, :, ::-1]  # 由R、G、B更改为B、G、R
    x = x[:, :, 1:3:-1]  # 由R、G、B更改为R、B、G
    # 改变一个数组的格式[需要改变的数组；新格式]
    x = np.reshape(x, (5,3,5))
    # 扩展数据
    x.repeat(100, axis=0)
    # 在某一维度，叠加numpy,或增加数据
    np.stack([x,x], axis=0)
    np.append(x, [[1, 2]], axis=0)
    # 经过不规则的slice操作，会改变数组的连续性
    # 将一个内存不连续存储的数组转换为内存连续存储的数组
    _ = np.ascontiguousarray(img)
import re
def regular_expression():
    """
    正则表达式
    :return:
    """
    # r表示字符串为转义的原始字符串
    # （.*） 第一个匹配分组，.*代表匹配除换行符之外的所有字符
    # （.*？） 第二个匹配分组，.*?多个问号，代表非贪婪模式，也就是说匹配符合条件的最少字符
    #  .* 没有括号包围，不计入匹配结果中,即不储存
    st = r'(.*) are (.*?) .*'
    # re.match只匹配字符串的开始，如果字符串开始不符合正则表达式，则匹配失败，函数返回None；
    # 而re.search匹配整个字符串，直到找到一个匹配。如果匹配规则可能是多种情况，则可能出现多种情况的第一个匹配结果。
    # $匹配字符串的结尾
    matchObj = re.search(r'<div>(.*)</div>$', "aa<div>test1</div>bb<div>test2</div>cc")
    if matchObj:
        print("matchObj.group() : ", matchObj.group())
        print("matchObj.group(1) : ", matchObj.group(1))   # 匹配式中第一个()的内容
    else:
        print("No match!!")

def opencv_note():
    """
    opencv的学习
    :return:
    """
    """
    opencv的数据类型为uint8,其他数据类型float64会自动转化，其他会报错
    """
    x = None
    # 读取图像[地址；何种模式读取图片，如灰度，彩色]
    img = cv2.imread(r"E:\py-workspace\test.jpg", cv2.IMREAD_GRAYSCALE)
    cv2.imshow("image", img)  # 显示图像，[窗口名，数据矩阵]
    cv2.imwrite("lena2.jpg", img)  # 保存图像
    # 在一个给定的时间内(单位ms)等待用户按键触发,0表示程序会无限制的等待用户的按键事件
    cv2.waitKey(20000)
    # 创建窗口，[窗口名；窗口尺寸]
    cv2.namedWindow('image', cv2.WINDOW_NORMAL)
    # 修改窗口大小【窗口名；宽；高】
    cv2.resizeWindow('image', 960, 540)
    # 关闭所有图像
    cv2.destroyAllWindows()
    # 关闭特定的窗口名
    cv2.destroyWindow("image")
    # 读取视频，进行灰度处理和垂直翻转后保存。
    cap = cv2.VideoCapture("E:\py-workspace\\test.avi")  # 读取视频, 0则读取本地内置摄像头
    cap.release()
    # window系统常用格式: DIVX(avi格式), mp4v(mp4格式)
    fourcc = cv2.VideoWriter_fourcc(*'DIVX')  # 编码器，DIVX是Windows系统的常用格式
    fps = cap.get(cv2.CAP_PROP_FPS)  # 获取原始视频的频帧
    # 获取原始视频的宽和高
    size = (int(cap.get(cv2.CAP_PROP_FRAME_WIDTH)),
            int(cap.get(cv2.CAP_PROP_FRAME_HEIGHT)))
    # 创建VideoWriter的对象，[输出名，编码器，频帧，大小，True彩色False灰色]
    out = cv2.VideoWriter('output.avi', fourcc, fps, size, False)
    out.release()
    # 设置视频的起始帧
    cap.set(cv2.CAP_PROP_POS_FRAMES, 2000)
    while cap.isOpened():  # 是否读取到
        ret, frame = cap.read()  # 是否读取到图片，每一帧的图片（mat）
        if ret is True:
            cv2.imshow("image", frame)
            cap.get(cv2.CAP_PROP_POS_FRAMES) # 获取视频的当前帧数
            gray = cv2.cvtColor(frame, cv2.COLOR_BGR2GRAY)  # 颜色转换
            frame = cv2.flip(gray, 0)  # 图像翻转，[mat；1水平翻转0垂直翻转-1水平垂直翻转]
            out.write(frame)  # 写入VideoWriter
            # 捕获返回的ascii值，与0xFF（1111 1111）是因为返回值只有后8位有效
            keycode = cv2.waitKey(1) & 0xFF
            if keycode == ord("q"):
                print("quit success")
                break
        else:
            break

    # 绘制不同几何图形，直线，矩形，圆，三角形，椭圆，多边形，添加文字
    # 绘制直线，[图像；起点；终点；颜色；粗细；类型]
    cv2.line(img, (0, 0), (511, 511), (255, 0, 0), 5)
    # 绘制椭圆，[图像；椭圆中心；长宽；旋转角度；绘制的起始角度（顺时针方向）；
    # 绘制的终止角度，如果是0，360，就是整个椭圆；颜色；粗细；类型 ]
    cv2.ellipse(img, (256, 256), (100, 50), 0, 0, 180, 255, -1, 8)
    # 绘制矩形[图像；两对角坐标；颜色；粗细]
    cv2.rectangle(img, (384, 0), (510, 128), (0, 255, 0), thickness=3)
    # 绘制圈，[图像；中心点坐标；半径；颜色；粗细；类型；小数点位数]
    cv2.circle(img, (447, 63), 63, (0, 0, 255), -1)
    # 绘制多边形，[图像；顶点坐标数组；是否闭合（首尾相连）；颜色]
    pts = np.array([[10, 5], [20, 30], [70, 20], [50, 10]], np.int32)
    pts2 = np.array([[100, 50], [200, 300], [700, 200], [500, 100]], np.int32)
    cv2.polylines(img, [pts, pts2], True, (0, 255, 255))
    # 添加文字[图片；文字内容；位置；字体类型；字体大小；颜色；粗细；线条类型]
    cv2.putText(img, 'OpenCV', (10, 500), cv2.FONT_HERSHEY_SIMPLEX, 4, (255, 255, 255), thickness=2, lineType=cv2.LINE_AA)
    # 读取写入XML, YML, JSON文件
    fs = cv2.FileStorage("abc.xml", cv2.FileStorage_READ) # filestorage_read读取，写入filestorage_wirte
    fs.write("mat1", np.random.randint(0, 255, [2, 2]))
    mat1 = fs.getNode("mat1").mat()  # 数组mat(),数字real(),字符串string()
    fs.release()
    # 边缘检测，[图像；minVal;maxVal;L2gradient求梯度大小的方程]
    edges = cv2.Canny(img, 100, 200)
    # 转为灰度图像
    img = cv2.cvtColor(img, cv2.COLOR_BGR2GRAY)
    # 高斯滤波是一种线性平滑滤波,能有效去除高斯噪声[图像；高斯矩阵大小；x方向标准差；y方向标准差]
    blur = cv2.GaussianBlur(img, (9, 9), sigmaX=2, sigmaY=2)
    # 模板匹配[待搜索图像；模板图像；计算匹配程度的方法]
    # 计算匹配程度的方法：
    # CV_TM_SQDIFF 平方差匹配法：该方法采用平方差来进行匹配；最好的匹配值为0；匹配越差，匹配值越大。
    # CV_TM_CCORR 相关匹配法：该方法采用乘法操作；数值越大表明匹配程度越好。
    # CV_TM_CCOEFF 相关系数匹配法：1表示完美的匹配；-1表示最差的匹配。
    # CV_TM_SQDIFF_NORMED 归一化平方差匹配法　
    # CV_TM_CCORR_NORMED 归一化相关匹配法　
    # CV_TM_CCOEFF_NORMED 归一化相关系数匹配法
    # result：
    # It must be single-channel 32-bit floating-point. If image is W*H and templ is w*h,
    # then result is (W-w+1)*(H-h+1)
    result = cv2.matchTemplate(img, img, cv2.TM_CCOEFF_NORMED)
    # 利用霍夫变换进行圆环检测【灰度图像；检测方法，目前唯一实现的方法；累加器分辨率与图像分辨率的反比；
    # 检测到圆心之间的最小距离；param1，Canny边缘函数的高阈值；param2，圆心检测阈值；
    # minRadius能检测到的最小圆半径；maxRadius，能检测到的最大圆半径】
    circles = cv2.HoughCircles(img, cv2.HOUGH_GRADIENT, 1, img.shape[0]/8, param1=200, param2=100, minRadius=0, maxRadius=0)
    # 角点检测【float32类型的图像；角点检测中考虑的领域大小；求导中使用的窗口大小；角点检测方程中的自由参数[0.04,0.06]】
    src_dst = cv2.cornerHarris(img, 2, 3, 0.04)
    # 角点检测【图像；检测的角度数量；最低角点质量；关键点距离】
    cv2.goodFeaturesToTrack(gray, 25, 0.01, 10)
    # 函数是一种形态学变化函数。数学形态学可以理解为一种滤波行为[图片；变化方式；方框大小]
    # cv2.MORPH_OPEN 进行开运算，先腐蚀后膨胀的过程。开运算可以用来消除小黑点
    # cv2.MORPH_CLOSE 先膨胀后腐蚀的过程。闭运算可以用来排除小黑洞。
    cv2.morphologyEx(img, cv2.MORPH_OPEN, kernel=np.ones((5, 5), np.uint8))
    # 固定阈值二值化 [图像；阈值；超过阈值所赋的值；二值化类型]，return:[得到图像的阈值[阈值设为0时，使用Otsu’s能自我选择阈值cv2.THRESH_OTSU]；图像]
    ret, dst = cv2.threshold(img, 127, 255, cv2.THRESH_OTSU)
    # 轮廓检测[图像；轮廓检索模式；轮廓的近似办法]
    # 轮廓检索模式：
    # cv2.RETR_EXTERNAL表示只检测外轮廓，
    # cv2.RETR_LIST检测的轮廓不建立等级关系。
    # cv2.RETR_CCOMP建立两个等级的轮廓，上面的一层为外边界，里面的一层为内孔的边界信息。如果内孔内还有一个连通物体，这个物体的边界也在顶层。
    # cv2.RETR_TREE建立一个等级树结构的轮廓；
    # 轮廓的逼近方法；
    # cv2.CHAIN_APPROX_NONE存储所有的轮廓点，cv2.CHAIN_APPROX_SIMPLE压缩水平方向，
    # cv2.CHAIN_APPROX_TC89_L1和cv2.CHAIN_APPROX_TC89_KCOS都是使用teh-Chinl chain近似算法]
    # return:[轮廓本身list(n)；每条轮廓对应的属性]
    # hierarchy:nparry(n*4),4:后一个轮廓、前一个轮廓、父轮廓、内嵌轮廓的索引编号
    contours, hierarchy = cv2.findContours(img, cv2.RETR_EXTERNAL, cv2.CHAIN_APPROX_SIMPLE)
    # 绘制轮廓，【图像；轮廓坐标点；0绘制外围轮廓，-1绘制所有轮廓；颜色】
    cv2.drawContours(img, contours, 0, (0, 0, 255), thickness=2)
    # 计算面积大小[np.array(n*2，多边形的顶点,np.int32类型)]
    area_points = np.array([[959.0, 13.0], [987.0, 16.0], [987.0, 20.0], [959.0, 17.0]], np.int32)
    area = cv2.contourArea(area_points)
    # 修改图像大小【图像；输出图像大小；dst输出图像；
    # fx(fy)缩放比例，如果它是0，那么它就会按照(double)dsize.width(height)/src.cols(rows)来计算;
    # interpolation指定插值的方式,重新计算像素的方式,INTER_LINEAR双线性插值（默认）】
    cv2.resize(img, (1080,1920), dst=None, fx=None, fy=None, interpolation=cv2.INTER_LINEAR)
    # 获取最小外接矩阵【点集坐标】， 返回【中心点坐标；（宽度，高度）；旋转角度（单位°）[-90,0]】
    # 旋转角度θ是水平轴（x轴）逆时针旋转，与碰到的矩形的第一条边的夹角。并且这个边的边长是width，另一条边边长是height。也就是说，在这里，width与height不是按照长短来定义的。
    rect = cv2.minAreaRect(pts)
    # 获取矩形的四个坐标点
    cv2.boxPoints(rect)
    # 设置鼠标的响应
    def on_Mouse(event, x, y, flags, param):
        """
        鼠标的回调函数
        :param event: event是 CV_EVENT_*变量之一
        Event:
            EVENT_MOUSEMOVE 0             //滑动
            EVENT_LBUTTONDOWN 1           //左键点击
            EVENT_RBUTTONDOWN 2           //右键点击
            EVENT_MBUTTONDOWN 3           //中键点击
            EVENT_LBUTTONUP 4             //左键放开
            EVENT_RBUTTONUP 5             //右键放开
            EVENT_MBUTTONUP 6             //中键放开
            EVENT_LBUTTONDBLCLK 7         //左键双击
            EVENT_RBUTTONDBLCLK 8         //右键双击
            EVENT_MBUTTONDBLCLK 9         //中键双击
        :param x:代表鼠标位于窗口的（x，y）坐标位置
        :param y:代表鼠标位于窗口的（x，y）坐标位置
        :param flags:  代表鼠标的拖拽事件，以及键盘鼠标联合事件，共有32种事件：
        flags:
            EVENT_FLAG_LBUTTON 1       //左鍵拖曳
            EVENT_FLAG_RBUTTON 2       //右鍵拖曳
            EVENT_FLAG_MBUTTON 4       //中鍵拖曳
            EVENT_FLAG_CTRLKEY 8       //(8~15)按Ctrl不放事件
            EVENT_FLAG_SHIFTKEY 16     //(16~31)按Shift不放事件
            EVENT_FLAG_ALTKEY 32       //(32~39)按Alt不放事件
        :param param:函数指针 标识了所响应的事件函数，相当于自定义了一个OnMouseAction()函数的ID。
        :return:
        """
        if event == cv2.EVENT_LBUTTONDOWN:
            print((x, y))
    cv2.setMouseCallback("images", on_Mouse, param=None)  # 为窗口添加鼠标的响应[窗口名；回调函数；传递回调函数的参数]

def pytorch_method():
    img = np.zeros((1920, 1080, 3), dtype=np.uint8)  # BGR格式,cv2读取标准格式

    device = torch.device('cuda:0')  # 定义计算使用的设备
    model = torch.load("yolov5l.pt", map_location=device)["model"]  # 加载模型
    _ = model.float()  # half>>>float，半浮点型转为浮点型
    img = torch.from_numpy(img).to(device).float()  # numpy>>>Tensor(uint8)>>>Tensor(float)
    _ = img.ndimension()  # 返回tensor的维度
    _ = img.unsqueeze(0)  # 在0处增加一个维度，与img.squeeze(0)相反
    _ = img.squeeze(0) # 删除在0处的维度
    torch.cuda.is_available()  # 查看cuda是否可用。
    torch.cuda.get_arch_list()  # 查看支持的compute capability
    # torch.cat将两个张量拼接在一起[待连接的张量;dim选择的维度]
    a, b = torch.ones((115, 6)), torch.zeros((50, 6))
    torch.cat((a, b), dim=0)
    # torch.nonzero取出不为0的索引[tensor;as_tuple输出是每一个维度都有一个一维的张量或每一行为非零元素的索引]
    a = torch.tensor([[1, 2], [3, 4], [5, 6]])
    torch.nonzero(a > 3, as_tuple=False)
    # tensor.view同ndarray.resize一样.
    a.view(-1)
    # tensor.max获取最大值及其索引[选择的维度;是否保持原有维度]
    value, index = a.max(1, keepdim=True)
    # torch.min一一比较两个tensor的最小值，后面的维度要相同
    torch.min(a, b)
    # tensor.prod元素的乘积[选择的维度]
    a.prod(1)
    # torch.mm两个张量的乘积
    torch.mm(a, b)
def tensorflow_method():
    tf.test.is_gpu_available()  # 测试GPU是否可用
