import cv2
import dlib
import os
import face_recognition
import numpy as np
# 定义已知的人脸知识图库
first_image = face_recognition.load_image_file("./data/hujt.jpg")
first_face_encoding = face_recognition.face_encodings(first_image)[0]
# 创建已知人名及对应的照片编码数组
known_face_encodings = [
    first_face_encoding
]
known_face_names = [
    "first"
]

detector = dlib.get_frontal_face_detector()  # 定义dlib库的检测算子
def forward(images):
    # 初始化一些变量
    face_locations = []
    face_names = []
    face_names_distances = []
    zoom_value = 2

    img  = images.copy()
    img = cv2.resize(img, (0, 0), fx=1/zoom_value, fy=1/zoom_value)  # 对照片进行缩放
    # gray_img = cv2.cvtColor(img, cv2.COLOR_BGR2GRAY)  # 将照片转为灰度图
    img = img[:, :, ::-1]# 将图片从BGR格式转化为RGB格式
    # 使用dlib检测算子，定位人脸坐标，并保存到数组中
    dets = detector(img, 1)
    for i, d in enumerate(dets):
        x1 = d.left() if d.left() > 0 else 0
        y1 = d.top() if d.top() > 0 else 0
        x2 = d.right() if d.right() > 0 else 0
        y2 = d.bottom() if d.bottom() > 0 else 0
        face_locations.append((y1,x2,y2,x1))
    # 先通过dlib.shape_predictor检测人脸68个特征点，再使用dlib提供的深度残差网络（ResNet）转换成128维的人脸特征向量。
    face_encodings = face_recognition.face_encodings(img, face_locations)
    # 循环检测到的人脸，并对比已知的人脸知识图库
    for face_encoding in face_encodings:
        name = "unknown"
        # 通过SVM、KNN算法等分类器，计算两张图片128维特征向量的欧式距离
        # tolerance为两张图片的相似性，通过调整它来改变识别的精确度，tolerance越低，精确度越高
        matches = face_recognition.compare_faces(known_face_encodings, face_encoding, tolerance=0.6)
        # 获取当前图片与图库中所有图片的欧氏距离
        face_distances = face_recognition.face_distance(known_face_encodings, face_encoding)
        # 获取欧式距离最小的那个，及相似度最大的
        best_match_index = np.argmin(face_distances)
        if matches[best_match_index]:
            name = known_face_names[best_match_index]
        face_names.append(name)
        distance = face_distances[best_match_index]*100
        face_names_distances.append("%.2f" % distance)
    # 将定位框及人名显示到图片中
    for (top, right, bottom, left), distance in zip(face_locations, face_names_distances):
        top *= zoom_value
        right *= zoom_value
        bottom *= zoom_value
        left *= zoom_value
        cv2.rectangle(images, (left, top), (right, bottom), (0, 0, 255), 2)
        font = cv2.FONT_HERSHEY_DUPLEX
        cv2.putText(images, distance, (left + 6, top - 6), font, 1.0, (255, 255, 255), 1)
    return [images, ",".join(face_names), ",".join(face_names_distances)]

if __name__ == "__main__":
    cap = cv2.VideoCapture("./test.avi")
    while cap.isOpened():
        ret, frame = cap.read()  # 是否读取到图片，每一帧的图片（mat）
        if ret is True:
            frame = forward(frame)
        else:
            break
        cv2.imshow('img', frame[0])
        # Hit 'q' on the keyboard to quit!
        if cv2.waitKey(1) & 0xFF == ord('q'):
            break
    cap.release()
    cv2.destroyAllWindows()