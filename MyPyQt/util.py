import traceback

from PyQt5.QtWidgets import QSplashScreen, QMessageBox
from PyQt5.QtGui import QPixmap


def image_wait(func):
    """
    等待界面的装饰器
    :param func: 需要等待完成的程序
    :return:
    """
    def wrapper():
        load_img = QPixmap('./image/wait.jpg')
        ratio = 200 / load_img.size().height()
        width = load_img.size().width() * ratio
        load_img = load_img.scaled(200, int(width))
        splash = QSplashScreen(load_img)
        splash.show()
        splash.showMessage('')
        try:
            res = func()
        except:
            splash.close()
            msg_box = QMessageBox(QMessageBox.Warning, '运行出错', traceback.format_exc())
            msg_box.exec_()
        else:
            splash.close()
            msg_box = QMessageBox(QMessageBox.Information, '提示', '已完成')
            msg_box.exec_()
            return res
    return wrapper