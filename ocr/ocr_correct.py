import sys
import os
# 解决pyqt打包成exe报ImportError: unable to find Qt5Core.dll on PATH。原因：pyqt5库对系统变量的加载存在bug
# if hasattr(sys, 'frozen'):
#     os.environ['PATH'] = sys._MEIPASS + ";" + os.environ['PATH']
import cgitb

from PyQt5 import QtWidgets
from PyQt5.QtWidgets import QFileDialog, QApplication
from PyQt5.QtGui import QPixmap, QImage, QRegExpValidator
from PyQt5.QtCore import Qt, QRegExp, QSize
import cv2
import numpy as np
from glob import glob
import pandas as pd
from identify_page import Ui_MainWindow as ui_identify_page
from main_page import Ui_MainWindow as ui_main_window
cgitb.enable(format="text")  # 解决pyqt异常，程序就崩溃，而没有任何提示的问题
# *******************************
# **********   主界面
# *******************************
class MyApp(QtWidgets.QMainWindow, ui_main_window):
    def __init__(self):
        QtWidgets.QMainWindow.__init__(self)
        ui_main_window.__init__(self)
        self.setupUi(self)

        self.setWindowTitle("Main")   # 设置窗口标题
        # 按钮绑定选择文件夹方法
        self.input_path.clicked.connect(self.chose_path)


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

    # 选择数据文件夹地址
    def chose_path(self):
        input_path = QFileDialog.getExistingDirectory(self, "选择输入文件夹地址")
        if not input_path:
            return
        self.hide()
        identify_page.show()
        identify_page.init_identify_page(input_path)

# *******************************
# **********   子页面
# *******************************
class IdentifyPage(QtWidgets.QMainWindow, ui_identify_page):
    jpg_path_list = None
    labels = None
    input_path = ""
    index = 0
    operation_msg = ""
    def __init__(self):
        QtWidgets.QMainWindow.__init__(self)
        ui_identify_page.__init__(self)
        self.setupUi(self)
        # 按钮绑定返回主界面的方法
        self.return_main.clicked.connect(self.returnMain)
        # 按钮绑定下一张的方法
        self.rc_next.clicked.connect(self.next_and_save)
        self.rc_last.clicked.connect(self.last_and_save)
        self.rc_del.clicked.connect(self.del_and_save)
        self.rc_jump.clicked.connect(self.jump_index)
        # 对输入框进行符号限制
        self.rc_label.setValidator(QRegExpValidator(QRegExp("[a-zA-Z0-9.|]+"), self.rc_label))
        self.rc_index.setValidator(QRegExpValidator(QRegExp("[0-9]+"), self.rc_index))

    # 关闭窗口的提示
    def closeEvent(self, event):
        box = QtWidgets.QMessageBox()
        Messages_S = "是否退出程序？"
        Messages_S = Messages_S
        reply = box.warning(self, '提示', Messages_S, box.Yes | box.No, box.Yes)
        if reply == box.Yes:
            self.labels.to_csv(os.path.join(self.input_path, "label.txt"), header=0, index=True,
                               encoding="gbk", sep=" ")
            event.accept()
        else:
            event.ignore()

    # 返回主界面
    def returnMain(self):
        self.hide()
        my_app.show()

    # 将cv2读取到的图片(numpy格式)转换成QPixmap格式，并显示到页面
    def show_picture(self, jpg_path):
        if jpg_path is None:
            self.pic_label.setPixmap(QPixmap(""))
            return
        img = QImage(jpg_path)
        ratio = np.array([img.height()/600, img.width()/400])
        max_ratio = ratio.max()
        mgnWidth = int(img.height() / max_ratio)
        mgnHeight = int(img.width() / max_ratio)
        size = QSize(mgnWidth, mgnHeight)
        jpg = QPixmap.fromImage(img.scaled(size, Qt.IgnoreAspectRatio))
        self.pic_label.resize(mgnWidth, mgnHeight)
        self.pic_label.setPixmap(jpg)

    # 初始化展示
    def init_identify_page(self, input_path):
        # 初始化参数
        self.input_path = input_path
        self.jpg_path_list = glob(os.path.join(self.input_path, "*.jpg"))
        label_path = os.path.join(self.input_path, "label.txt")
        if not os.path.exists(label_path):
            data_dict = {"name": self.jpg_path_list, "label": [""] * len(self.jpg_path_list)}
            data = pd.DataFrame(data_dict)
            data = data.astype("str")
            data["name"] = data["name"].apply(lambda x: os.path.basename(x))
            data.to_csv(os.path.join(self.input_path, "label.txt"), index=False, header=False, encoding="gbk", sep=" ")
        self.labels = pd.read_csv(label_path, header=None, sep=" ",
                                  index_col=0, encoding="gbk", keep_default_na=False)
        self.labels[1] = self.labels[1].astype("str")
        self.index = 0
        # 展示到页面
        self.show_message()

    # 根据所给索引，将图片和标签展示出来
    def show_message(self):
        jpg_path = self.jpg_path_list[self.index]
        name = os.path.basename(jpg_path)
        if name not in self.labels.index:
            label = ""
            self.labels = self.labels.append(pd.Series([label], index=[1], name=name))
        else:
            label = self.labels.loc[name][1]
        self.rc_name.setText(name)
        self.show_picture(jpg_path)
        self.rc_label.setText(label)
        self.rc_label.selectAll()
        remain_num = len(self.jpg_path_list) - self.index - 1
        self.statusBar.showMessage('%s索引：%d，剩余：%d' % (self.operation_msg, self.index, remain_num))

    # 保存并显示下一张图片
    def next_and_save(self):
        if self.index < len(self.jpg_path_list) - 1:
            self.change_and_save()
            self.index += 1
            self.show_message()
        else:
            return

    # 保存并显示上一张图片
    def last_and_save(self):
        if self.index > 0:
            self.change_and_save()
            self.index -= 1
            self.show_message()
        else:
            return

    # 删除并保存
    def del_and_save(self):
        self.rc_label.setText("")
        self.change_and_save()
        if self.index < len(self.jpg_path_list) - 1:
            self.index += 1
            self.show_message()
        else:
            return

    # 跳转到给定位置
    def jump_index(self):
        jump_index = int(self.rc_index.text())
        if -1<jump_index<len(self.jpg_path_list):
            self.index = jump_index
            self.show_message()
    # 修改标签并保存
    def change_and_save(self):
        label = self.rc_label.text()
        name = os.path.basename(self.jpg_path_list[self.index])

        if label == "":
            self.operation_msg = '删除%s，' % name
        elif self.labels.loc[name][1] == label:
            self.operation_msg = '未修改%s，' % name
            return
        else:
            self.operation_msg = '修改%s，' % name
        self.labels.loc[name][1] = label
        self.labels.to_csv(os.path.join(self.input_path, "label.txt"), header=0, index=True,
                           encoding="gbk", sep=" ")


    # 检测键盘操作
    def keyPressEvent(self, event):
        # 这里event.key（）显示的是按键的编码
        # print("按下：" + str(event.key()))
        if event.key() == Qt.Key_Up:
            self.last_and_save()
        if event.key() == Qt.Key_Enter or event.key() == Qt.Key_Return:
            self.next_and_save()
        if event.key() == Qt.Key_Control:
            self.del_and_save()


if __name__ == "__main__":
    app = QtWidgets.QApplication(sys.argv)
    my_app = MyApp()  # 实例化主页面
    identify_page= IdentifyPage()  # 实例化识别界面
    my_app.show()  # 显示主页面
    sys.exit(app.exec_())