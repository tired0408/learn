import cgitb
from PyQt5 import uic, QtWidgets
import sys
from PyQt5.QtWidgets import QFileDialog
from PyQt5.QtGui import QPixmap, QImage
from PyQt5.QtCore import Qt, QThread, pyqtSignal
from face_recoge import forward
import cv2
import os

cgitb.enable(format="text")  # 解决pyqt异常，程序就崩溃，而没有任何提示的问题

ui_main_window, _ = uic.loadUiType("./main.ui")  # 加载主页面的样式
ui_identify_page, _ = uic.loadUiType("./identify_page.ui")  # 加载图片模式的页面样式

ui_progress, _ = uic.loadUiType("./progress.ui")  # 加载进度条窗口的样式

# *******************************
# **********   主界面
# *******************************
class MyApp(QtWidgets.QMainWindow, ui_main_window):
    def __init__(self):
        QtWidgets.QMainWindow.__init__(self)
        ui_main_window.__init__(self)
        self.setupUi(self)

        self.setWindowTitle("Main")   # 设置窗口标题
        # 按钮绑定进入子页面的方法
        self.picture_model.clicked.connect(lambda: self.enterIdentifyPage("picture"))
        self.camera_model.clicked.connect(lambda: self.enterIdentifyPage("camera"))
        self.accuracy.clicked.connect(lambda: self.enterIdentifyPage("accuracy"))

    # 关闭窗口的提示
    def closeEvent(self, event):
        box = QtWidgets.QMessageBox()
        Messages_S = "是否退出程序？"
        Messages_S = Messages_S
        reply = box.warning(self, '提示', Messages_S, box.Yes | box.No, box.Yes)
        if reply == box.Yes:
            event.accept()
        else:
            event.ignore()

    # 根据所选模式进入相应界面
    def enterIdentifyPage(self, model_type):
        self.hide()
        if model_type == "camera":
            camera_model.show()
        elif model_type == "picture":
            picture_model.show()
        elif model_type == "accuracy":
            accuracy_model.show()
# *******************************
# **********   子页面
# *******************************
class IdentifyPage(QtWidgets.QMainWindow, ui_identify_page):

    def __init__(self):
        QtWidgets.QMainWindow.__init__(self)
        ui_identify_page.__init__(self)

        self.setupUi(self)

        # 按钮绑定返回主界面的方法
        self.return_main.clicked.connect(self.returnMain)

    # 关闭窗口的提示
    def closeEvent(self, event):
        box = QtWidgets.QMessageBox()
        Messages_S = "是否退出程序？"
        Messages_S = Messages_S
        reply = box.warning(self, '提示', Messages_S, box.Yes | box.No, box.Yes)
        if reply == box.Yes:
            event.accept()
        else:
            event.ignore()

    # 返回主界面
    def returnMain(self):
        self.hide()
        my_app.show()

    # 将cv2读取到的图片(numpy格式)转换成QPixmap格式，并显示到页面
    def show_picture(self, jpg_numpy):
        jpg_numpy = cv2.resize(jpg_numpy, (600, 400), interpolation=cv2.INTER_LINEAR)
        jpg = cv2.cvtColor(jpg_numpy, cv2.COLOR_BGR2RGB)
        jpg = QImage(jpg.data, jpg.shape[1], jpg.shape[0], QImage.Format_RGB888)
        jpg = QPixmap(QPixmap.fromImage(jpg))
        self.pic_label.setPixmap(jpg)

    # 展示预测结果
    def show_recognition_result(self, name, precision):
        self.name.setText(name)
        self.precision.setText(precision)

    # 清除图片
    def close_picture(self):
        self.name.setText("")  # 清空姓名的值
        self.precision.setText("")  # 清空置信度的值
        self.pic_label.setPixmap(QPixmap(""))  # 移除图片


# *******************************
# ********** 图像模式
# *******************************
class PictureModel(IdentifyPage):
    save_image = None
    def __init__(self):
        IdentifyPage.__init__(self)
        self.setWindowTitle("图片模式")
        self.open.setText("导入图片")
        self.open.clicked.connect(self.import_picture)
        self.close.setText("清除图片")
        self.close.clicked.connect(self.close_picture)
        self.start.setText("识别图片")
        self.start.clicked.connect(self.identify_picture)
        self.stop.hide()
        self.accuracy_group.hide()

    # 导入图片的功能
    def import_picture(self):
        filenames = QFileDialog.getOpenFileName(self, "选择图片", "", "jpg files(*.jpg)")
        if not filenames[0]:
            return
        # 清空姓名、置信度的值
        self.name.setText("")
        self.precision.setText("")
        # 用cv2读取图片，并保存该图片
        jpg = cv2.imread(filenames[0])
        self.save_image = jpg
        self.show_picture(jpg)

    # 调用父类清除图片的方法，并清理缓存
    def close_picture(self):
        super().close_picture()
        self.save_image = None

    # 识别图片
    def identify_picture(self):
        if self.save_image is not None:
            result = forward(self.save_image)
            self.show_picture(result[0])
            self.show_recognition_result(result[1], result[2])


# *******************************
# ********** 摄像头模式
# *******************************
class CameraModel(IdentifyPage):

    def __init__(self):
        IdentifyPage.__init__(self)
        self.setWindowTitle("摄像头模式")
        self.open.setText("打开摄像头")
        self.open.clicked.connect(self.open_camera)
        self.close.setText("关闭摄像头")
        self.close.clicked.connect(self.close_camera)
        self.start.setText("开始检测")
        self.start.clicked.connect(self.start_identify)
        self.stop.setText("停止检测")
        self.stop.clicked.connect(self.stop_identify)
        self.accuracy_group.hide()
        # 连接子进程的信号和槽函数
        self.cameraThread = CameraThread()
        self.cameraThread.signal.connect(self.show_camera)

    # 打开摄像头
    def open_camera(self):
        self.cameraThread.is_run = True
        self.cameraThread.start()
        progress_window.show()

    # 关闭摄像头
    def close_camera(self):
        self.cameraThread.is_run = False

    # 显示摄像头内容
    def show_camera(self, content_list):
        if content_list[0] is False:
            self.close_picture()
        else:
            jpg = content_list[0]
            name = content_list[1]
            precision = content_list[2]
            progress_window.close()
            self.show_picture(jpg)
            self.show_recognition_result(name, precision)

    # 开始检测
    def start_identify(self):
        self.cameraThread.is_identify = True

    # 停止检测
    def stop_identify(self):
        self.cameraThread.is_identify = False
# *******************************
# ********** 摄像头模式处理线程
# *******************************
class CameraThread(QThread):
    signal = pyqtSignal(list)
    is_run = True
    is_identify = False
    def __init__(self, parent=None):
        super(CameraThread, self).__init__(parent)

    def run(self):
        cap = cv2.VideoCapture("./test.avi")
        while cap.isOpened():  # 是否读取到
            ret, frame = cap.read()  # 是否读取到图片，每一帧的图片（mat）
            if ret is False:
                break
            if not self.is_run:
                self.signal.emit([False])
                break
            if self.is_identify:
                self.signal.emit(forward(frame))
            else:
                self.signal.emit([frame, "", ""])

# *******************************
# ********** 统计精准度模式
# *******************************
class AccuracyModel(IdentifyPage):

    images_list = []
    current_image_index = 0
    def __init__(self):
        IdentifyPage.__init__(self)
        self.setWindowTitle("统计精准度模式")
        self.open.setText("选择文件夹")
        self.open.clicked.connect(self.selectFiles)
        self.close.setText("上一张")
        self.close.clicked.connect(self.last_picture)
        self.start.setText("下一张")
        self.start.clicked.connect(self.next_picture)
        self.stop.hide()

        # 连接子进程的信号和槽函数
        self.accuracyThread = AccuracyThread()
        self.accuracyThread.signal.connect(self.deal_thread)

    # 选择文件夹
    def selectFiles(self):
        input_path = QFileDialog.getExistingDirectory(self, "选择输入文件夹地址")
        if not input_path:
            return
        self.accuracy.setText("")
        self.accuracyThread.input_path = input_path
        self.accuracyThread.start()
        progress_window.label.setText("统计中，请稍等...")
        progress_window.show()

    # 关闭摄像头
    def close_camera(self):
        self.cameraThread.is_run = False

    # 处理线程的情况
    def deal_thread(self, accuracy):
        progress_window.close()
        self.show_picture(self.images_list[self.current_image_index][0])
        self.show_recognition_result(self.images_list[self.current_image_index][1],
                                     self.images_list[self.current_image_index][2])
        self.accuracy.setText("%.2f" % accuracy)

    # 上一张
    def last_picture(self):
        if self.current_image_index > 0:
            self.current_image_index = self.current_image_index - 1
        self.show_picture(self.images_list[self.current_image_index][0])
        self.show_recognition_result(self.images_list[self.current_image_index][1],
                                     self.images_list[self.current_image_index][2])

    # 下一张
    def next_picture(self):
        if self.current_image_index + 1 < len(self.images_list):
            self.current_image_index = self.current_image_index + 1
        self.show_picture(self.images_list[self.current_image_index][0])
        self.show_recognition_result(self.images_list[self.current_image_index][1],
                                     self.images_list[self.current_image_index][2])
# *******************************
# ********** 统计精准度模式处理线程
# *******************************
class AccuracyThread(QThread):
    statistical_name = "first"
    input_path = None
    signal = pyqtSignal(float)
    def __init__(self, parent=None):
        super(AccuracyThread, self).__init__(parent)

    def run(self):
        picture_name_list = os.listdir(self.input_path)
        total = len(picture_name_list)
        ture_num = 0
        for picture_name in picture_name_list:
            jpg = cv2.imread(os.path.join(self.input_path, picture_name))
            result = forward(jpg)
            if result[1] == self.statistical_name:
                ture_num += 1
            accuracy_model.images_list.append(result)
        self.signal.emit(ture_num*100/total)
# *******************************
# **********   进度条窗口
# *******************************
class ProgressDialog(QtWidgets.QDialog, ui_progress):
    def __init__(self):
        QtWidgets.QDialog.__init__(self)
        ui_progress.__init__(self)
        self.setupUi(self)
        # 设置标题（在使用了QT Designer 设计好以后可以这么使用）
        self.setWindowTitle('打开摄像头中')

    # **********   关闭窗口的提示
    def closeEvent(self, event):
        event.accept()


if __name__ == "__main__":
    app = QtWidgets.QApplication(sys.argv)
    my_app = MyApp()  # 实例化主页面
    picture_model= PictureModel()  # 实例化图片模式界面
    camera_model = CameraModel()  # 实例化摄像头模式界面
    accuracy_model = AccuracyModel()  # 实例化统计精准度模式页面
    progress_window = ProgressDialog()  # 实例化进度条窗口
    my_app.show()  # 显示主页面

    sys.exit(app.exec_())