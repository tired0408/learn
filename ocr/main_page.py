# -*- coding: utf-8 -*-

# Form implementation generated from reading ui file 'main.ui'
#
# Created by: PyQt5 UI code generator 5.13.2
#
# WARNING! All changes made in this file will be lost!


from PyQt5 import QtCore, QtGui, QtWidgets


class Ui_MainWindow(object):
    def setupUi(self, MainWindow):
        MainWindow.setObjectName("MainWindow")
        MainWindow.resize(409, 224)
        MainWindow.setMinimumSize(QtCore.QSize(409, 224))
        MainWindow.setMaximumSize(QtCore.QSize(409, 224))
        self.centralwidget = QtWidgets.QWidget(MainWindow)
        self.centralwidget.setObjectName("centralwidget")
        self.groupBox_3 = QtWidgets.QGroupBox(self.centralwidget)
        self.groupBox_3.setGeometry(QtCore.QRect(20, 120, 371, 81))
        self.groupBox_3.setToolTipDuration(9)
        self.groupBox_3.setObjectName("groupBox_3")
        self.gridLayout = QtWidgets.QGridLayout(self.groupBox_3)
        self.gridLayout.setObjectName("gridLayout")
        self.input_path = QtWidgets.QPushButton(self.groupBox_3)
        self.input_path.setMinimumSize(QtCore.QSize(181, 51))
        font = QtGui.QFont()
        font.setFamily("微软雅黑")
        font.setPointSize(14)
        font.setBold(True)
        font.setWeight(75)
        self.input_path.setFont(font)
        self.input_path.setObjectName("input_path")
        self.gridLayout.addWidget(self.input_path, 0, 0, 1, 1)
        self.label_2 = QtWidgets.QLabel(self.centralwidget)
        self.label_2.setGeometry(QtCore.QRect(20, 30, 391, 81))
        font = QtGui.QFont()
        font.setFamily("微软雅黑 Light")
        font.setPointSize(45)
        font.setBold(True)
        font.setWeight(75)
        self.label_2.setFont(font)
        self.label_2.setObjectName("label_2")
        MainWindow.setCentralWidget(self.centralwidget)

        self.retranslateUi(MainWindow)
        QtCore.QMetaObject.connectSlotsByName(MainWindow)

    def retranslateUi(self, MainWindow):
        _translate = QtCore.QCoreApplication.translate
        MainWindow.setWindowTitle(_translate("MainWindow", "MainWindow"))
        self.groupBox_3.setTitle(_translate("MainWindow", "GroupBox"))
        self.input_path.setText(_translate("MainWindow", "选择文件夹"))
        self.label_2.setText(_translate("MainWindow", "文字识别纠正"))
