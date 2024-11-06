from PyQt5 import QtCore, QtGui, uic, QtWidgets
import sys
import cgitb
import cv2
import os
import numpy as np
import time
cgitb.enable(format="text")  # 解决pyqt异常，程序就崩溃，而没有任何提示的问题

from util import image_wait
"""
获取主界面类:
1. 直接加载ui文件
2. 使用: pyuic5 -o 目标文件名.py 源文件名.ui , 转成py文件后导入
"""
ui_path = os.path.join(os.path.dirname(__file__), "./GUI/main.ui")
UiMainWindow, _ = uic.loadUiType(ui_path)


class MyApp(QtWidgets.QMainWindow, UiMainWindow):
    def __init__(self):
        QtWidgets.QMainWindow.__init__(self)
        UiMainWindow.__init__(self)
        self.setupUi(self)
        # 设置标题（在使用了QT Designer 设计好以后可以这么使用）
        self.setWindowTitle('HujtMain')

        self.btn_1.clicked.connect(self.set_status_value)  # 修改状态栏
        self.btn_2.clicked.connect(lambda : self.show_picture(np.zeros((500, 500, 3))))  # 显示图片
        # 线程相关操作
        self.thread_one = ThreadOne()   # 创建线程实例
        self.thread_one.signal_box.connect(self.deal_signal)  # 为线程信号槽绑定处理函数
        self.btn_3.clicked.connect(self.thread_one.start)  # 启动线程
        self.btn_4.clicked.connect(self.thread_one.stop)  # 停止线程
        self.btn_5.clicked.connect(self.thread_one.pause)  # 暂停线程
        self.btn_6.clicked.connect(self.thread_one.resume)  # 启动线程
        # 往listWidget中添加元素
        self.btn_7.clicked.connect(self.add_item_to_list)
        # 绑定列表右键删除菜单
        # 将ContextMenuPolicy设置为Qt.CustomContextMenu,否则无法使用customContextMenuRequested信号
        self.setContextMenuPolicy(QtCore.Qt.CustomContextMenu)
        self.contextMenu = QtWidgets.QMenu(self)  # 创建QMenu
        self.actionA = self.contextMenu.addAction(QtGui.QIcon("images/0.png"), u'|  删除')  # 添加按钮
        self.customContextMenuRequested.connect(self.show_context_menu)  # 显示菜单
        self.contextMenu.triggered[QtWidgets.QAction].connect(self.del_execute_func)  # 绑定删除函数
        self.btn_8.clicked.connect(self.calculate_time)  # 进度条计时
        self.btn_9.clicked.connect(self.choose_file)  # 选择文件
        self.btn_10.clicked.connect(self.choose_folder)  # 选择文件夹
        self.btn_11.clicked.connect(self.save_file)  # 保存文件
        self.btn_12.clicked.connect(self.image_wait_show)

    def set_status_value(self):
        """修改状态栏的值"""
        self.statusbar.showMessage("填写状态栏的值")

    def show_picture(self, img):
        """
        显示图片
        :param img: numpy的格式, shape(x,x,3)
        :return:
        """
        if img is None:
            self.pic_label.setPixmap(QtGui.QPixmap(""))
            return
        label_height = self.pic_label.height()
        label_width = self.pic_label.width()
        ratio = np.array([img.shape[1]/label_width, img.shape[0]/label_height])
        max_ratio = ratio.max()
        img = cv2.resize(img, (int(img.shape[1]/max_ratio), int(img.shape[0]/max_ratio)), interpolation=cv2.INTER_LINEAR)
        img = QtGui.QImage(img.data, img.shape[1], img.shape[0], QtGui.QImage.Format_RGB888)
        img = QtGui.QPixmap(QtGui.QPixmap.fromImage(img))
        self.pic_label.setPixmap(img)

    def deal_signal(self, msg):
        """获取线程类中信号槽的数据，并输出到QTableWidget中"""
        row = self.tableWidget.rowCount()
        if row > 4:
            self.tableWidget.setRowCount(0)
            row = 0
        self.tableWidget.insertRow(row)
        msg = QtWidgets.QTableWidgetItem(msg)
        msg.setTextAlignment(QtCore.Qt.AlignCenter)
        # 设置物件的状态为只可被选择（未设置可编辑）
        msg.setFlags(QtCore.Qt.ItemIsSelectable | QtCore.Qt.ItemIsEnabled)
        exec_time = QtWidgets.QTableWidgetItem(str(int(time.time())))
        exec_time.setTextAlignment(QtCore.Qt.AlignCenter)
        exec_time.setFlags(QtCore.Qt.ItemIsSelectable | QtCore.Qt.ItemIsEnabled)

        self.tableWidget.setItem(row, 0, msg)
        self.tableWidget.setItem(row, 1, exec_time)

    def add_item_to_list(self):
        """添加数据到列表中"""
        count = self.listWidget.count()
        if count > 4:
            self.listWidget.clear()
        self.listWidget.addItem("行数据%d" % int(time.time()))

    def show_context_menu(self):
        """显示右键菜单"""
        items = self.listWidget.selectedIndexes()
        if items:
            self.contextMenu.show()
            self.contextMenu.exec_(QtGui.QCursor.pos())

    def del_execute_func(self):
        """删除数据"""
        del_index = self.listWidget.selectedIndexes()[0].row()
        self.listWidget.removeItemWidget(self.listWidget.takeItem(del_index))

    def calculate_time(self):
        """计算已运行时间，以进度条形式展示"""
        for i in range(101):
            self.progressBar.setValue(i)
            time.sleep(0.01)

    def choose_file(self):
        """选择文件"""
        file_paths = QtWidgets.QFileDialog.getOpenFileName(self, "选择文件", "", "csv files(*.csv)")
        if not file_paths[0]:
            return
        print(file_paths)  # 进行文件处理

    def choose_folder(self):
        """选择文件夹"""
        directory = QtWidgets.QFileDialog.getExistingDirectory(self, "选取文件夹", "C:/")
        print(directory)

    def save_file(self):
        """保存文件"""

    @staticmethod
    @image_wait
    def image_wait_show():
        """等待界面显示"""
        time.sleep(1)


class ThreadOne(QtCore.QThread):
    """线程1"""
    signal_box = QtCore.pyqtSignal(str)  # 定义信号槽

    def __init__(self, parent=None):
        super(ThreadOne, self).__init__(parent)
        self._isRun = True
        self._isParse = False
        self.mutex = QtCore.QMutex()
        self.cond = QtCore.QWaitCondition()

    def run(self):
        self._isRun = True
        while self._isRun:
            self.mutex.lock()  # 线程锁
            if self._isParse:
                self.cond.wait(self.mutex)  # 进入等待
            self.signal_box.emit("Execute")  # 向信号槽中传递数据
            time.sleep(1)
            self.mutex.unlock()  # 线程解锁

    def stop(self):
        """停止线程"""
        self.signal_box.emit("Stop")
        self._isRun = False

    def pause(self):
        """暂停线程"""
        self.signal_box.emit("Pause")
        self._isParse = True

    def resume(self):
        """启动线程"""
        self.signal_box.emit("Start")
        self._isParse = False
        self.cond.wakeAll()


# 这段主程序创建了一个新的 Qt Gui 应用。，每个 QT 应用都可以通过命令行进行配置，所以必须传入sys.argv 参数。
if __name__ == "__main__":
    app = QtWidgets.QApplication(sys.argv)
    window = MyApp()  # 实例化主页面类
    window.show()   # 显示主页面
    sys.exit(app.exec_())