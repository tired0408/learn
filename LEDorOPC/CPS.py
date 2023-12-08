#!/usr/bin/python3
# --*-- coding:utf-8 --*--
import os
import glob
os.environ["CUDA_VISIBLE_DEVICES"]="2"

import numpy as np
import argparse
import random
import time
import cv2
import os
import sys, os

sys.path.append("./Files/")
import time
from time import sleep,ctime
import random
import numpy as np
from cv2 import *
import copy
from PyQt5 import QtCore, QtGui, uic, QtWidgets
import queue
import threading


from base64 import b64encode
from json import dumps



# 创建线程锁
QueueLock = threading.Lock()
# 创建Frame队列
FrameQueue = queue.Queue(5)
# 保存视频队列
SaveQueue = queue.Queue(1)
# ***********************************************
# ***********************************************
# 获取实时视频流线程，即使销毁无用帧
# ***********************************************
# ***********************************************
class Get_Frame (threading.Thread):   #继承父类threading.Thread
    def __init__(self, threadID, RSTP_Path):
        threading.Thread.__init__(self)
        self.threadID = threadID
        self.RSTP_Path = RSTP_Path

        self.Save_Flag = False
        self.Exit_Flag = False

    def run(self): #把要执行的代码写到run函数里面 线程在创建后会直接运行run函数
        print('[INFO] Get_Frame Start ...')
        video_capture = cv2.VideoCapture(self.RSTP_Path)
        while True:
            ret, frame = video_capture.read()
            if ret != True:
                break

            if self.Save_Flag:
                # 保存文件
                filestr = 'abcdefghijklmnopqrstuvwxyz'
                fileend = ''
                for m in range(5):
                    fileend = fileend + random.choice(filestr)
                time_str = time.strftime('%Y-%m-%d_%H-%M-%S', time.localtime(time.time()))
                File_Name_1 = './Log_Pic/' + time_str + '_' + fileend + '_' + 'frame' + '.jpg'
                # 保存图片
                cv2.imwrite(File_Name_1, frame)
                self.Save_Flag = False

            QueueLock.acquire()
            if FrameQueue.qsize() > 2:
                FrameQueue.get()
            else:
                FrameQueue.put([ret,frame])
            QueueLock.release()

            # 获取状态队列,保存视频或者退出软件用
            if not SaveQueue.empty():
                SaveQueue_Get = SaveQueue.get()
                self.Save_Flag = SaveQueue_Get[0]
                self.Exit_Flag = SaveQueue_Get[1]
                # self.labelname = SaveQueue_Get[2]
                # self.points = SaveQueue_Get[3]

        # Release everything if job is finished
        video_capture.release()

# ***********************************************
# ***********************************************
# 执行动作
# ***********************************************
# ***********************************************
'''class Vision_Do(threading.Thread):  # 继承父类threading.Thread
    def __init__(self, threadID, lane, CM, DU):
        threading.Thread.__init__(self)
        self.threadID = threadID
        self.lane = lane
        self.CM = CM
        self.DU = DU


    def run(self):  # 把要执行的代码写到run函数里面 线程在创建后会直接运行run函数

        Zoom_Value = 0.25
        # ********************
        # 一些判断规则，当前车道，或者所有车道等等
        # ********************
        Value_List = [{'Pass_Position': 1080, 'Adjuct_X': 10, 'Min_X': 1440, 'Max_X': 1952},
                      {'Pass_Position': 1080, 'Adjuct_X': 10, 'Min_X': 1848, 'Max_X': 2372},
                      {'Pass_Position': 1080, 'Adjuct_X': 10, 'Min_X': 2256, 'Max_X': 2744},
                      {'Pass_Position': 1080, 'Adjuct_X': 10, 'Min_X': 2648, 'Max_X': 3108},
                      {'Pass_Position': 1080, 'Adjuct_X': 10, 'Min_X': 3016, 'Max_X': 3452},
                      {'Pass_Position': 1080, 'Adjuct_X': 10, 'Min_X': 1440, 'Max_X': 3552}]

        # 需要停下的坐标像素值
        Pass_Position = Value_List[self.lane - 1]['Pass_Position'] * Zoom_Value
        # 这里需要做换算！！！！，意思是在这个误差范围内的都可以认为通过
        WUCHA_CM = int(self.CM)

        Min_X = int(Value_List[self.lane - 1]['Min_X'] * Zoom_Value)
        Max_X = int(Value_List[self.lane - 1]['Max_X'] * Zoom_Value)

        cv2.namedWindow("Window_1", 0)
        cv2.resizeWindow("Window_1", 960, 540)
        cv2.namedWindow("Window_3", 0)
        cv2.resizeWindow("Window_3", 960, 540)
        cv2.ocl.setUseOpenCL(False)

        savetimeflag = time.time()
        savepositionY = 0

        # 图片识别
        LABELS = ['container','truck']

        # 加载用于可视化给定实例分割的颜色集合

        COLORS = ['0,0,255', '255,0,0', '0,255,0', '255,255,0', '0,255,255', '255,255,255']
        COLORS = [np.array(c.split(",")).astype("int") for c in COLORS]
        COLORS = np.array(COLORS, dtype="uint8")

        # Mask R-CNN 权重路径及模型配置文件
        weightsPath = './output/frozen_inference_graph.pb'
        configPath = './output/mask_rcnn.pbtxt'

        # 加载预训练的 Mask R-CNN 模型(90 classes)
        print("[INFO] loading Mask R-CNN from disk...")
        net = cv2.dnn.readNetFromTensorflow(weightsPath, configPath)

        # ********************
        # 主程序循环
        # ********************
        while True:

            if not FrameQueue.empty():
                # 队列线程锁
                QueueLock.acquire()
                FrameQueue_Get = FrameQueue.get()
                QueueLock.release()
                ret = FrameQueue_Get[0]
                frame = FrameQueue_Get[1]


                # frame = cv2.imread('10.jpg')
                # 读取图片
                image = cv2.resize(frame, None, fx=Zoom_Value, fy=Zoom_Value, interpolation=cv2.INTER_CUBIC)
                (H, W) = image.shape[:2]
                clone = image.copy()
                image = cv2.UMat(image)
                # 构建输入图片 blob
                blob = cv2.dnn.blobFromImage(image, swapRB=True, crop=False)
                net.setInput(blob)

                start = time.time()
                # forward 计算，输出图片中目标的边界框坐标以及每个目标的像素级分割
                (boxes, masks) = net.forward(["detection_out_final", "detection_masks"])
                # print(masks)
                end = time.time()

                # Mask R-CNN 的时间统计
                # print("[INFO] Mask R-CNN took {:.6f} seconds".format(end - start))
                # print("[INFO] boxes shape: {}".format(boxes.shape))
                # print("[INFO] masks shape: {}".format(masks.shape))
                # print(boxes.shape[2])
                # loop over the number of detected objects

                for i in range(0, boxes.shape[2]):
                    # 检测的 class ID 及对应的置信度(概率)
                    classID = int(boxes[0, 0, i, 1])
                    confidence = boxes[0, 0, i, 2]

                    # 过滤低置信度预测结果
                    if confidence > 0.4:
                        # 用于可视化

                        # 将边界框坐标缩放回相对于图片的尺寸，然后计算边界框的width和height
                        box = boxes[0, 0, i, 3:7] * np.array([W, H, W, H])
                        (startX, startY, endX, endY) = box.astype("int")
                        boxW = endX - startX
                        boxH = endY - startY

                        # 提取目标的像素级分割
                        mask = masks[i, classID]
                        print(11111,mask.shape[:2],boxW,boxH)
                        # resize mask以保持与边界框的维度一致
                        mask = cv2.resize(mask, (boxW, boxH),
                                          interpolation=cv2.INTER_CUBIC)
                        # 根据设定阈值，得到二值化mask.
                        mask = (mask > 0.1)

                        # 提取图片的 ROI
                        roi = clone[startY:endY, startX:endX]

                        # 可视化
                        if startX > Min_X:
                            # 将二值mask转换为:0和255
                            visMask = (mask * 255).astype("uint8")
                            instance = cv2.bitwise_and(roi, roi, mask=visMask)
                            cv2.imshow("masks", visMask)
                            # 可视化提取的 ROI、mask 以及对应的分割实例
                            # cv2.imshow("ROI", roi)
                            # cv2.imshow("Mask", visMask)
                            # cv2.imshow("Segmented", instance)

                            # 只提取 ROI 的 masked 区域
                            roi = roi[mask]

                            # 随机选择一种颜色，用于可视化特定的实例分割
                            color = random.choice(COLORS)
                            # 通过融合选择的颜色和 ROI 进行融合，创建透明覆盖图
                            blended = ((0.4 * color) + (0.6 * roi)).astype("uint8")

                            # 替换原始图片的融合 ROI 区域
                            clone[startY:endY, startX:endX][mask] = blended

                            # 画出图片中实例的边界框
                            color = [int(c) for c in color]
                            cv2.rectangle(clone, (startX, startY), (endX, endY), (0,0,0), 5)

                            # 画出预测的类别标签以及对应的实例概率
                            text = "{}: {:.4f}".format(LABELS[classID], confidence)
                            cv2.putText(clone, text, (startX, startY + 15),
                                        cv2.FONT_HERSHEY_SIMPLEX, 0.5, color, 2)

                            if (abs(startY - savepositionY) > 150) or ((time.time() - savetimeflag) > 30):
                                # 保存图片
                                # SaveQueue.put([True, False])
                                savepositionY = startY
                                savetimeflag = time.time()



                # show
                cv2.imshow("Window_1", clone)

                Press_Key = cv2.waitKey(1) & 0xFF
                if Press_Key == ord('a'):
                    SaveQueue.put([True, False])
                    print('[INFO] 开始保存图片...')
                elif Press_Key == ord('b'):
                    SaveQueue.put([False, False])
                    print('[INFO] 暂停保存图片...')
                elif Press_Key == ord('q'):
                    SaveQueue.put([False, True])
                    break


def RUN():
    T1 = Get_Frame(1,"rtsp://admin:xmhf12345@172.16.149.200:554")
    # T1 = Get_Frame(1, "./avi/output.avi")
    T2 = Vision_Do(2, 6, 10, 1)

    T1.start()
    T2.start()

    T1.join()
    T2.join()  # 线程守护，保证每个线程都运行完成

    print('over %s' % ctime())

RUN()'''

# 下面是可以用的！！！！！！！
# 下面是可以用的！！！！！！！
# 下面是可以用的！！！！！！！
# 下面是可以用的！！！！！！！
# 下面是可以用的！！！！！！！
# 下面是可以用的！！！！！！！
# 下面是可以用的！！！！！！！
# 下面是可以用的！！！！！！！
# 下面是可以用的！！！！！！！

# ***********************************************
# ***********************************************
# 执行动作
# ***********************************************
# ***********************************************
class Vision_Do(threading.Thread):  # 继承父类threading.Thread
    def __init__(self, threadID, lane, CM, DU):
        threading.Thread.__init__(self)
        self.threadID = threadID
        self.lane = lane
        self.CM = CM
        self.DU = DU



    def run(self):  # 把要执行的代码写到run函数里面 线程在创建后会直接运行run函数
        import tensorflow as tf
        Zoom_Value = 0.25
        # ********************
        # 一些判断规则，当前车道，或者所有车道等等
        # ********************
        Value_List = [{'Pass_Position': 1080, 'Adjuct_X': 10, 'Min_X': 1440, 'Max_X': 1952},
                      {'Pass_Position': 1080, 'Adjuct_X': 10, 'Min_X': 1848, 'Max_X': 2372},
                      {'Pass_Position': 1080, 'Adjuct_X': 10, 'Min_X': 2256, 'Max_X': 2744},
                      {'Pass_Position': 1080, 'Adjuct_X': 10, 'Min_X': 2648, 'Max_X': 3108},
                      {'Pass_Position': 1080, 'Adjuct_X': 10, 'Min_X': 3016, 'Max_X': 3452},
                      {'Pass_Position': 1080, 'Adjuct_X': 10, 'Min_X': 1440, 'Max_X': 3552}]

        # 需要停下的坐标像素值
        Pass_Position = Value_List[self.lane - 1]['Pass_Position'] * Zoom_Value
        # 这里需要做换算！！！！，意思是在这个误差范围内的都可以认为通过
        WUCHA_CM = int(self.CM)

        Min_X = int(Value_List[self.lane - 1]['Min_X'] * Zoom_Value)
        Max_X = int(Value_List[self.lane - 1]['Max_X'] * Zoom_Value)

        cv2.namedWindow("Window_1", 0)
        cv2.resizeWindow("Window_1", 960, 540)
        cv2.namedWindow("Window_3", 0)
        cv2.resizeWindow("Window_3", 960, 540)

        savetimeflag = time.time()
        savepositionY = 0

        # 图片识别
        LABELS = ['container','truck']

        # 加载用于可视化给定实例分割的颜色集合

        COLORS = ['0,0,255', '255,0,0', '0,255,0', '255,255,0', '0,255,255', '255,255,255']
        COLORS = [np.array(c.split(",")).astype("int") for c in COLORS]
        COLORS = np.array(COLORS, dtype="uint8")

        # Mask R-CNN 权重路径及模型配置文件
        weightsPath = './output/frozen_inference_graph.pb'
        configPath = './output/mask_rcnn.pbtxt'

        PATH_TO_CKPT = './output/frozen_inference_graph.pb'

        # Load a (frozen) Tensorflow model into memory
        detection_graph = tf.Graph()
        with detection_graph.as_default():
            od_graph_def = tf.GraphDef()
            with tf.gfile.GFile(PATH_TO_CKPT, 'rb') as fid:
                serialized_graph = fid.read()
                od_graph_def.ParseFromString(serialized_graph)
                tf.import_graph_def(od_graph_def, name='')

        image_tensor = detection_graph.get_tensor_by_name('image_tensor:0')
        # Each box represents a part of the image where a particular
        # object was detected.
        gboxes = detection_graph.get_tensor_by_name('detection_boxes:0')
        # Each score represent how level of confidence for each of the objects.
        # Score is shown on the result image, together with the class label.
        gscores = detection_graph.get_tensor_by_name('detection_scores:0')
        gclasses = detection_graph.get_tensor_by_name('detection_classes:0')
        gnum_detections = detection_graph.get_tensor_by_name('num_detections:0')
        masks_detections = detection_graph.get_tensor_by_name('detection_masks:0')


        # TODO: Add class names showing in the image
        def detect_image_objects(image, sess, detection_graph):
            # Expand dimensions since the model expects images to have
            # shape: [1, None, None, 3]
            image_np_expanded = np.expand_dims(image, axis=0)

            # Actual detection.

            (boxes, scores, classes, num_detections, masks) = sess.run(
                [gboxes, gscores, gclasses, gnum_detections, masks_detections],
                feed_dict={image_tensor: image_np_expanded})

            # Visualization of the results of a detection.
            boxes = np.squeeze(boxes)
            scores = np.squeeze(scores)
            height, width = image.shape[:2]
            for i in range(boxes.shape[0]):
                if (scores is None or
                        scores[i] > 0.5):
                    ymin, xmin, ymax, xmax = boxes[i]
                    ymin = int(ymin * height)
                    ymax = int(ymax * height)
                    xmin = int(xmin * width)
                    xmax = int(xmax * width)

                    score = None if scores is None else scores[i]
                    font = cv2.FONT_HERSHEY_SIMPLEX
                    text_x = np.max((0, xmin - 10))
                    text_y = np.max((0, ymin - 10))
                    cv2.putText(image, 'Detection score: ' + str(score),
                                (text_x, text_y), font, 0.4, (0, 255, 0))
                    cv2.rectangle(image, (xmin, ymin), (xmax, ymax),
                                  (0, 255, 0), 1)
            return image


        with detection_graph.as_default():
            with tf.Session(graph=detection_graph) as sess:
                # video_path = './3.avi'
                # capture = cv2.VideoCapture(video_path)
                while True:
                    if cv2.waitKey(30) & 0xFF == ord('q'):
                        break
                    if not FrameQueue.empty():
                        # 队列线程锁
                        QueueLock.acquire()
                        FrameQueue_Get = FrameQueue.get()
                        QueueLock.release()
                        ret = FrameQueue_Get[0]
                        frame = FrameQueue_Get[1]
                        frame = imread('1.png')
                        frame = cv2.resize(frame, None, fx=Zoom_Value, fy=Zoom_Value, interpolation=cv2.INTER_CUBIC)

                        t_start = time.clock()
                        detect_image_objects(frame, sess, detection_graph)
                        t_end = time.clock()
                        print('detect time per frame: ', t_end - t_start)
                        cv2.imshow('Window_1', frame)

                cv2.destroyAllWindows()


def RUN():
    # T1 = Get_Frame(1,"rtsp://admin:xmhf12345@172.16.149.200:554")
    T1 = Get_Frame(1, "./avi/output.avi")
    T2 = Vision_Do(2, 6, 10, 1)

    T1.start()
    T2.start()

    T1.join()
    T2.join()  # 线程守护，保证每个线程都运行完成

    print('over %s' % ctime())

RUN()
"""
Created on Sat Nov  4 15:05:09 2017

@author: shirhe-lyh
"""
'''import time

import cv2
import numpy as np
import tensorflow as tf

# --------------Model preparation----------------
# Path to frozen detection graph. This is the actual model that is used for
# the object detection.
PATH_TO_CKPT = './output/frozen_inference_graph.pb'

# Load a (frozen) Tensorflow model into memory
detection_graph = tf.Graph()
with detection_graph.as_default():
	od_graph_def = tf.GraphDef()
	with tf.gfile.GFile(PATH_TO_CKPT, 'rb') as fid:
		serialized_graph = fid.read()
		od_graph_def.ParseFromString(serialized_graph)
		tf.import_graph_def(od_graph_def, name='')

image_tensor = detection_graph.get_tensor_by_name('image_tensor:0')
# Each box represents a part of the image where a particular
# object was detected.
gboxes = detection_graph.get_tensor_by_name('detection_boxes:0')
# Each score represent how level of confidence for each of the objects.
# Score is shown on the result image, together with the class label.
gscores = detection_graph.get_tensor_by_name('detection_scores:0')
gclasses = detection_graph.get_tensor_by_name('detection_classes:0')
gnum_detections = detection_graph.get_tensor_by_name('num_detections:0')


# TODO: Add class names showing in the image
def detect_image_objects(image, sess, detection_graph):
	# Expand dimensions since the model expects images to have
	# shape: [1, None, None, 3]
	image_np_expanded = np.expand_dims(image, axis=0)

	# Actual detection.
	(boxes, scores, classes, num_detections) = sess.run(
		[gboxes, gscores, gclasses, gnum_detections],
		feed_dict={image_tensor: image_np_expanded})

	# Visualization of the results of a detection.
	boxes = np.squeeze(boxes)
	scores = np.squeeze(scores)
	height, width = image.shape[:2]
	for i in range(boxes.shape[0]):
		if (scores is None or
				scores[i] > 0.5):
			ymin, xmin, ymax, xmax = boxes[i]
			ymin = int(ymin * height)
			ymax = int(ymax * height)
			xmin = int(xmin * width)
			xmax = int(xmax * width)

			score = None if scores is None else scores[i]
			font = cv2.FONT_HERSHEY_SIMPLEX
			text_x = np.max((0, xmin - 10))
			text_y = np.max((0, ymin - 10))
			cv2.putText(image, 'Detection score: ' + str(score),
						(text_x, text_y), font, 0.4, (0, 255, 0))
			cv2.rectangle(image, (xmin, ymin), (xmax, ymax),
						  (0, 255, 0), 2)
	return image


with detection_graph.as_default():
	with tf.Session(graph=detection_graph) as sess:
		video_path = './3.avi'
		capture = cv2.VideoCapture(video_path)
		while capture.isOpened():
			if cv2.waitKey(30) & 0xFF == ord('q'):
				break
			ret, frame = capture.read()
			if not ret:
				break

			t_start = time.clock()
			detect_image_objects(frame, sess, detection_graph)
			t_end = time.clock()
			print('detect time per frame: ', t_end - t_start)
			cv2.imshow('detected', frame)
		capture.release()
		cv2.destroyAllWindows()'''