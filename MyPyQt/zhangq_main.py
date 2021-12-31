# coding=gbk
import sys,os
import time

from PyQt5 import QtCore, QtGui, uic, QtWidgets

from PyQt5.QtCore import *
from PyQt5.QtGui import *
from PyQt5.QtWidgets import *

from PIL import Image, ImageGrab, ImageFilter

# 导入跨文件的全局变量
import zhangq_global_var as globalvar

import cgitb
cgitb.enable(format="text")  # 解决pyqt异常，程序就崩溃，而没有任何提示的问题

# 一个主界面、两个绘画界面、一个进度条界面
Ui_MainWindow, _ = uic.loadUiType(".\GUI\zhangq\MainWindow.ui")
Ui_Draw_0, QtBaseClass_1 = uic.loadUiType(".\GUI\zhangq\Draw.ui")
Ui_Draw_1, QtBaseClass_2 = uic.loadUiType(".\GUI\zhangq\Draw_Num.ui")
Ui_Progress, QtBaseClass_3 = uic.loadUiType(".\GUI\zhangq\Progress.ui")

# 跨文件全局变量初始化,用不到再删除
globalvar._init()

# *******************************
# **********   主界面
# *******************************
class MyApp(QtWidgets.QMainWindow, Ui_MainWindow):
    def __init__(self):
        QtWidgets.QMainWindow.__init__(self)
        Ui_MainWindow.__init__(self)
        self.setupUi(self)
        # 设置标题（在使用了QT Designer 设计好以后可以这么使用）
        self.setWindowTitle('Main')
        # 给按钮定义点击函数
        self.Single_Button.clicked.connect(self.closeWindow)
        self.Num_Button.clicked.connect(self.Num_Vision)
        self.Draw_0_Button.clicked.connect(self.OpenDraw0)
        self.Draw_1_Button.clicked.connect(self.OpenDraw1)
        # 创建子进程实例
        self.single_thread = SingleVisionThread()
        self.NumThread = Num_Vision_Thread()
        self.Draw0_window = Draw0_Dialog()
        self.Draw1_window = Draw1_Dialog()
        # 给线程的信号槽绑定处理函数
        self.single_thread.finish_signal.connect(self.finish_single)
        self.NumThread.Finish_Signal.connect(self.finish_single)
        self.Draw0_window.Finish_Signal.connect(self.Draw0_Finish_Single)
        self.Draw1_window.finish_signal.connect(self.Draw1_Finish_Single)

    #设置显示图片
    def showpic(self,filename):
        #调用QtGui.QPixmap方法，打开一个图片，存放在变量png中
        png=QPixmap(filename)
        Position = self.pic_label.geometry()
        scaredPixmap = png.scaled(Position.width()+Position.left(),
                                  Position.height()+ Position.top(),
                                  aspectRatioMode=Qt.KeepAspectRatio)
        self.pic_label.setPixmap(scaredPixmap)

    # 关闭主窗口操作
    def closeWindow(self):
        self.close()

    # 关闭窗口的提示
    def closeEvent(self, event):
        box = QtWidgets.QMessageBox()
        Messages_S = "确认退出？"
        Messages_S = Messages_S
        reply = box.warning(self, '提示', Messages_S, box.Yes | box.No, box.Yes)
        if reply == box.Yes:
            event.accept()
        else:
            event.ignore()

    # 打开警告弹窗
    def OpenWarningBox(self, Messages_S):
        box = QtWidgets.QMessageBox()
        Messages_S = Messages_S
        box.warning(self, '提示', Messages_S, box.Ok)

    # 按钮按下对应的操作
    def single_vision(self):
        self.textEdit.setText('')
        self.Num_Button.setDisabled(True)
        self.Draw_0_Button.setDisabled(True)
        self.Draw_1_Button.setDisabled(True)
        self.single_thread.start()
        Progress_window.show()

    # 结束以后需要关闭进度条，显示结果，并且启用按钮
    def finish_single(self, Result):
        Progress_window.close()
        self.textEdit.setText(Result)
        self.Num_Button.setDisabled(False)
        self.Draw_0_Button.setDisabled(False)
        self.Draw_1_Button.setDisabled(False)

    # 对应第二个按钮的功能
    def Num_Vision(self):
        self.textEdit.setText('')
        filenames = QFileDialog.getOpenFileName(self, "选择图片", "", "jpg files(*.jpg)")
        if len(filenames[0]) > 0:
            filename = filenames[0]
            # 设置全局变量，供别的文件访问
            globalvar.Set_value('filename', filename)
            # print(Visions.knn_Single(filename))
            self.NumThread.start()
            Progress_window.show()
            self.showpic(filename)

    # 对应手写识别的两个窗口
    def OpenDraw0(self):
        self.textEdit.setText('')
        self.Draw0_window.pos_xy = []
        self.Draw0_window.show()
        # self.hide()
    def OpenDraw1(self):
        self.textEdit.setText('')
        self.Draw1_window.pos_xy = []
        self.Draw1_window.show()
        # self.hide()

    # 手写识别窗口关闭进行的动作
    def Draw0_Finish_Single(self):
        filename = './cache.jpg'
        globalvar.Set_value('filename', filename)
        # print(Visions.knn_Single(filename))
        self.single_thread.start()
        Progress_window.show()
        self.showpic(filename)
    def Draw1_Finish_Single(self):
        filename = './cache.jpg'
        # 设置跨文件全局变量，供其它程序访问
        globalvar.Set_value('filename', filename)
        # 开启子线程
        self.NumThread.start()
        # 打开进度窗口
        Progress_window.show()
        # 显示该图片
        self.showpic(filename)



# *******************************
# **********   单个字符手写界面
# *******************************
class Draw0_Dialog(QtWidgets.QDialog, Ui_Draw_0):
    Finish_Signal = pyqtSignal()
    def __init__(self):
        QtWidgets.QDialog.__init__(self)
        Ui_Draw_0.__init__(self)
        self.setupUi(self)

        # 设置标题（在使用了QT Designer 设计好以后可以这么使用）
        self.setWindowTitle('单个字符手写识别')

        # setMouseTracking设置为False，否则不按下鼠标时也会跟踪鼠标事件
        self.setMouseTracking(False)

        self.pos_xy = []

    # 关闭窗口的提示
    def closeDialog(self):
        self.close()

    # **********   关闭窗口的提示
    def closeEvent(self, event):
        Position = self.geometry()
        img = ImageGrab.grab((Position.left(),
                              Position.top(),
                              Position.width() + Position.left(),
                              Position.height() + Position.top()))
        img = img.filter(ImageFilter.BLUR)
        img.save("./cache.jpg")
        # 设置全局变量，供别的文件访问
        event.accept()
        self.Finish_Signal.emit()


    def paintEvent(self, event):
        painter = QPainter()
        painter.begin(self)
        pen = QPen(Qt.black, 10, Qt.SolidLine)
        painter.setPen(pen)
        if len(self.pos_xy) > 1:
            point_start = self.pos_xy[0]
            for pos_tmp in self.pos_xy:
                point_end = pos_tmp

                if point_end == (-1, -1):
                    point_start = (-1, -1)
                    continue
                if point_start == (-1, -1):
                    point_start = point_end
                    continue

                painter.drawLine(point_start[0], point_start[1], point_end[0], point_end[1])
                point_start = point_end

        painter.end()

    def mouseMoveEvent(self, event):
        #中间变量pos_tmp提取当前点
        pos_tmp = (event.pos().x(), event.pos().y())
        #pos_tmp添加到self.pos_xy中
        self.pos_xy.append(pos_tmp)

        self.update()

    def mouseReleaseEvent(self, event):
        pos_test = (-1, -1)
        self.pos_xy.append(pos_test)

        self.update()

# *******************************
# **********   多个字符手写界面
# *******************************
class Draw1_Dialog(QtWidgets.QDialog, Ui_Draw_1):
    finish_signal = pyqtSignal()
    def __init__(self):
        QtWidgets.QDialog.__init__(self)
        Ui_Draw_1.__init__(self)
        self.setupUi(self)

        # 设置标题（在使用了QT Designer 设计好以后可以这么使用）
        self.setWindowTitle('多个字符手写识别')

        # setMouseTracking设置为False，否则不按下鼠标时也会跟踪鼠标事件
        self.setMouseTracking(False)

        self.pos_xy = []

    # 关闭窗口的提示
    def closeDialog(self):
        self.close()

    # **********   关闭窗口的提示
    def closeEvent(self, event):
        Position = self.geometry()
        img = ImageGrab.grab((Position.left(),
                              Position.top(),
                              Position.width() + Position.left(),
                              Position.height() + Position.top()))
        img = img.filter(ImageFilter.BLUR)
        img.save("./cache.jpg")
        # 设置全局变量，供别的文件访问
        event.accept()
        self.finish_signal.emit()


    def paintEvent(self, event):
        painter = QPainter()
        painter.begin(self)
        pen = QPen(Qt.black, 10, Qt.SolidLine)
        painter.setPen(pen)
        if len(self.pos_xy) > 1:
            point_start = self.pos_xy[0]
            for pos_tmp in self.pos_xy:
                point_end = pos_tmp

                if point_end == (-1, -1):
                    point_start = (-1, -1)
                    continue
                if point_start == (-1, -1):
                    point_start = point_end
                    continue

                painter.drawLine(point_start[0], point_start[1], point_end[0], point_end[1])
                point_start = point_end

        painter.end()

    def mouseMoveEvent(self, event):
        #中间变量pos_tmp提取当前点
        pos_tmp = (event.pos().x(), event.pos().y())
        #pos_tmp添加到self.pos_xy中
        self.pos_xy.append(pos_tmp)

        self.update()

    def mouseReleaseEvent(self, event):
        pos_test = (-1, -1)
        self.pos_xy.append(pos_test)

        self.update()

# *******************************
# **********   进度条窗口
# *******************************
class Progress_Dialog(QtWidgets.QDialog, Ui_Progress):
    def __init__(self):
        QtWidgets.QDialog.__init__(self)
        Ui_Progress.__init__(self)
        self.setupUi(self)
        # 设置标题（在使用了QT Designer 设计好以后可以这么使用）
        self.setWindowTitle('正在识别')

    # **********   关闭窗口的提示
    def closeEvent(self, event):
        # window.show()
        event.accept()


# *************************************
# ********** 多线程任务（识别单个字符）
# *************************************
class SingleVisionThread(QThread):
    # 扫描箱号后第一次更新显示
    finish_signal = pyqtSignal(str)
    def __init__(self, parent=None):
        super(SingleVisionThread, self).__init__(parent)

    def run(self):
        # 处理程序写到这里
        # 处理程序写到这里
        # 处理程序写到这里
        # 处理程序写到这里
        # 处理程序写到这里
        Result = '识别结束'
        time.sleep(1)
        self.finish_signal.emit(str(Result))

# *************************************
# ********** 多线程任务（识别多个字符）
# *************************************
class Num_Vision_Thread(QThread):
    # 扫描箱号后第一次更新显示
    Finish_Signal = pyqtSignal(str)
    def __init__(self, parent=None):
        super(Num_Vision_Thread, self).__init__(parent)

    def run(self):
        # 处理程序写到这里
        # 处理程序写到这里
        # 处理程序写到这里
        # 处理程序写到这里
        # 处理程序写到这里
        Result = '识别结束'
        time.sleep(5)
        self.Finish_Signal.emit(str(Result))




# 这段主程序创建了一个新的 Qt Gui 应用。，每个 QT 应用都可以通过命令行进行配置，所以必须传入sys.argv 参数。
if __name__ == "__main__":

    app = QtWidgets.QApplication(sys.argv)
    # Draw0_Window = Draw0_Dialog()
    # Draw1_Window = Draw1_Dialog()
    Progress_window = Progress_Dialog()
    window = MyApp()
    window.show()

    sys.exit(app.exec_())
