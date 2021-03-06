# -*- coding: utf-8 -*-

# Form implementation generated from reading ui file 'ui.ui'
#
# Created by: PyQt5 UI code generator 5.15.1
#
# WARNING: Any manual changes made to this file will be lost when pyuic5 is
# run again.  Do not edit this file unless you know what you are doing.


from PyQt5 import QtCore, QtGui, QtWidgets


class Ui_MainWindow(object):
    def setupUi(self, MainWindow):
        MainWindow.setObjectName("MainWindow")
        MainWindow.resize(692, 492)
        self.centralwidget = QtWidgets.QWidget(MainWindow)
        self.centralwidget.setObjectName("centralwidget")
        self.verticalLayout = QtWidgets.QVBoxLayout(self.centralwidget)
        self.verticalLayout.setSizeConstraint(QtWidgets.QLayout.SetDefaultConstraint)
        self.verticalLayout.setContentsMargins(3, 3, 3, 3)
        self.verticalLayout.setSpacing(0)
        self.verticalLayout.setObjectName("verticalLayout")
        self.horizontalLayout_2 = QtWidgets.QHBoxLayout()
        self.horizontalLayout_2.setObjectName("horizontalLayout_2")
        self.logo_label = QtWidgets.QLabel(self.centralwidget)
        sizePolicy = QtWidgets.QSizePolicy(QtWidgets.QSizePolicy.Fixed, QtWidgets.QSizePolicy.Fixed)
        sizePolicy.setHorizontalStretch(0)
        sizePolicy.setVerticalStretch(0)
        sizePolicy.setHeightForWidth(self.logo_label.sizePolicy().hasHeightForWidth())
        self.logo_label.setSizePolicy(sizePolicy)
        self.logo_label.setMaximumSize(QtCore.QSize(100, 100))
        self.logo_label.setText("")
        self.logo_label.setPixmap(QtGui.QPixmap("Logo.ico"))
        self.logo_label.setScaledContents(True)
        self.logo_label.setObjectName("logo_label")
        self.horizontalLayout_2.addWidget(self.logo_label)
        self.label = QtWidgets.QLabel(self.centralwidget)
        sizePolicy = QtWidgets.QSizePolicy(QtWidgets.QSizePolicy.Preferred, QtWidgets.QSizePolicy.Fixed)
        sizePolicy.setHorizontalStretch(0)
        sizePolicy.setVerticalStretch(0)
        sizePolicy.setHeightForWidth(self.label.sizePolicy().hasHeightForWidth())
        self.label.setSizePolicy(sizePolicy)
        font = QtGui.QFont()
        font.setFamily("IBM Plex Sans")
        font.setPointSize(32)
        self.label.setFont(font)
        self.label.setObjectName("label")
        self.horizontalLayout_2.addWidget(self.label)
        self.verticalLayout.addLayout(self.horizontalLayout_2)
        spacerItem = QtWidgets.QSpacerItem(20, 5, QtWidgets.QSizePolicy.Minimum, QtWidgets.QSizePolicy.Fixed)
        self.verticalLayout.addItem(spacerItem)
        self.line_2 = QtWidgets.QFrame(self.centralwidget)
        self.line_2.setFrameShape(QtWidgets.QFrame.HLine)
        self.line_2.setFrameShadow(QtWidgets.QFrame.Sunken)
        self.line_2.setObjectName("line_2")
        self.verticalLayout.addWidget(self.line_2)
        self.label_2 = QtWidgets.QLabel(self.centralwidget)
        sizePolicy = QtWidgets.QSizePolicy(QtWidgets.QSizePolicy.Preferred, QtWidgets.QSizePolicy.Fixed)
        sizePolicy.setHorizontalStretch(0)
        sizePolicy.setVerticalStretch(0)
        sizePolicy.setHeightForWidth(self.label_2.sizePolicy().hasHeightForWidth())
        self.label_2.setSizePolicy(sizePolicy)
        self.label_2.setMinimumSize(QtCore.QSize(0, 50))
        font = QtGui.QFont()
        font.setFamily("IBM Plex Sans")
        font.setPointSize(14)
        font.setBold(False)
        font.setWeight(50)
        self.label_2.setFont(font)
        self.label_2.setObjectName("label_2")
        self.verticalLayout.addWidget(self.label_2, 0, QtCore.Qt.AlignHCenter)
        self.privacy_policy_btn = QtWidgets.QPushButton(self.centralwidget)
        sizePolicy = QtWidgets.QSizePolicy(QtWidgets.QSizePolicy.MinimumExpanding, QtWidgets.QSizePolicy.Fixed)
        sizePolicy.setHorizontalStretch(0)
        sizePolicy.setVerticalStretch(0)
        sizePolicy.setHeightForWidth(self.privacy_policy_btn.sizePolicy().hasHeightForWidth())
        self.privacy_policy_btn.setSizePolicy(sizePolicy)
        self.privacy_policy_btn.setMinimumSize(QtCore.QSize(500, 40))
        self.privacy_policy_btn.setStyleSheet("QPushButton {\n"
"background-color: rgb(71, 118, 93);\n"
"color: rgb(254, 254, 254);\n"
"border-radius: 20px;\n"
"font: 57 14pt \"IBM Plex Sans\";\n"
"}")
        self.privacy_policy_btn.setAutoDefault(False)
        self.privacy_policy_btn.setObjectName("privacy_policy_btn")
        self.verticalLayout.addWidget(self.privacy_policy_btn, 0, QtCore.Qt.AlignHCenter)
        spacerItem1 = QtWidgets.QSpacerItem(20, 10, QtWidgets.QSizePolicy.Minimum, QtWidgets.QSizePolicy.Fixed)
        self.verticalLayout.addItem(spacerItem1)
        self.line = QtWidgets.QFrame(self.centralwidget)
        self.line.setFrameShape(QtWidgets.QFrame.HLine)
        self.line.setFrameShadow(QtWidgets.QFrame.Sunken)
        self.line.setObjectName("line")
        self.verticalLayout.addWidget(self.line)
        self.label_3 = QtWidgets.QLabel(self.centralwidget)
        sizePolicy = QtWidgets.QSizePolicy(QtWidgets.QSizePolicy.Preferred, QtWidgets.QSizePolicy.Fixed)
        sizePolicy.setHorizontalStretch(0)
        sizePolicy.setVerticalStretch(0)
        sizePolicy.setHeightForWidth(self.label_3.sizePolicy().hasHeightForWidth())
        self.label_3.setSizePolicy(sizePolicy)
        self.label_3.setMinimumSize(QtCore.QSize(0, 50))
        font = QtGui.QFont()
        font.setFamily("IBM Plex Sans")
        font.setPointSize(14)
        font.setBold(False)
        font.setWeight(50)
        self.label_3.setFont(font)
        self.label_3.setObjectName("label_3")
        self.verticalLayout.addWidget(self.label_3, 0, QtCore.Qt.AlignHCenter)
        self.horizontalLayout = QtWidgets.QHBoxLayout()
        self.horizontalLayout.setSizeConstraint(QtWidgets.QLayout.SetDefaultConstraint)
        self.horizontalLayout.setContentsMargins(-1, 0, -1, -1)
        self.horizontalLayout.setSpacing(6)
        self.horizontalLayout.setObjectName("horizontalLayout")
        self.path_le = QtWidgets.QLineEdit(self.centralwidget)
        self.path_le.setMinimumSize(QtCore.QSize(0, 45))
        font = QtGui.QFont()
        font.setPointSize(12)
        self.path_le.setFont(font)
        self.path_le.setText("")
        self.path_le.setObjectName("path_le")
        self.horizontalLayout.addWidget(self.path_le)
        self.change_path_btn = QtWidgets.QPushButton(self.centralwidget)
        sizePolicy = QtWidgets.QSizePolicy(QtWidgets.QSizePolicy.Fixed, QtWidgets.QSizePolicy.Fixed)
        sizePolicy.setHorizontalStretch(0)
        sizePolicy.setVerticalStretch(0)
        sizePolicy.setHeightForWidth(self.change_path_btn.sizePolicy().hasHeightForWidth())
        self.change_path_btn.setSizePolicy(sizePolicy)
        self.change_path_btn.setMinimumSize(QtCore.QSize(200, 40))
        self.change_path_btn.setStyleSheet("QPushButton {\n"
"background-color: rgb(71, 118, 93);\n"
"color: rgb(254, 254, 254);\n"
"border-radius: 20px;\n"
"font: 57 14pt \"IBM Plex Sans\";\n"
"}")
        self.change_path_btn.setAutoDefault(False)
        self.change_path_btn.setObjectName("change_path_btn")
        self.horizontalLayout.addWidget(self.change_path_btn)
        self.verticalLayout.addLayout(self.horizontalLayout)
        self.horizontalLayout_4 = QtWidgets.QHBoxLayout()
        self.horizontalLayout_4.setObjectName("horizontalLayout_4")
        self.label_4 = QtWidgets.QLabel(self.centralwidget)
        sizePolicy = QtWidgets.QSizePolicy(QtWidgets.QSizePolicy.Preferred, QtWidgets.QSizePolicy.Fixed)
        sizePolicy.setHorizontalStretch(0)
        sizePolicy.setVerticalStretch(0)
        sizePolicy.setHeightForWidth(self.label_4.sizePolicy().hasHeightForWidth())
        self.label_4.setSizePolicy(sizePolicy)
        self.label_4.setMinimumSize(QtCore.QSize(0, 50))
        font = QtGui.QFont()
        font.setFamily("IBM Plex Sans")
        font.setPointSize(14)
        font.setBold(False)
        font.setWeight(50)
        self.label_4.setFont(font)
        self.label_4.setObjectName("label_4")
        self.horizontalLayout_4.addWidget(self.label_4, 0, QtCore.Qt.AlignRight)
        spacerItem2 = QtWidgets.QSpacerItem(40, 20, QtWidgets.QSizePolicy.Fixed, QtWidgets.QSizePolicy.Minimum)
        self.horizontalLayout_4.addItem(spacerItem2)
        self.system = QtWidgets.QComboBox(self.centralwidget)
        font = QtGui.QFont()
        font.setFamily("IBM Plex Sans")
        font.setPointSize(14)
        self.system.setFont(font)
        self.system.setFrame(True)
        self.system.setObjectName("system")
        self.system.addItem("")
        self.system.addItem("")
        self.horizontalLayout_4.addWidget(self.system, 0, QtCore.Qt.AlignLeft)
        self.verticalLayout.addLayout(self.horizontalLayout_4)
        self.install_btn = QtWidgets.QPushButton(self.centralwidget)
        sizePolicy = QtWidgets.QSizePolicy(QtWidgets.QSizePolicy.Preferred, QtWidgets.QSizePolicy.Fixed)
        sizePolicy.setHorizontalStretch(0)
        sizePolicy.setVerticalStretch(0)
        sizePolicy.setHeightForWidth(self.install_btn.sizePolicy().hasHeightForWidth())
        self.install_btn.setSizePolicy(sizePolicy)
        self.install_btn.setMinimumSize(QtCore.QSize(250, 40))
        self.install_btn.setStyleSheet("QPushButton {\n"
"background-color: rgb(71, 118, 93);\n"
"color: rgb(254, 254, 254);\n"
"border-radius: 20px;\n"
"font: 57 14pt \"IBM Plex Sans\";\n"
"}")
        self.install_btn.setAutoDefault(False)
        self.install_btn.setObjectName("install_btn")
        self.verticalLayout.addWidget(self.install_btn, 0, QtCore.Qt.AlignHCenter)
        spacerItem3 = QtWidgets.QSpacerItem(20, 10, QtWidgets.QSizePolicy.Minimum, QtWidgets.QSizePolicy.Fixed)
        self.verticalLayout.addItem(spacerItem3)
        self.line_3 = QtWidgets.QFrame(self.centralwidget)
        self.line_3.setFrameShape(QtWidgets.QFrame.HLine)
        self.line_3.setFrameShadow(QtWidgets.QFrame.Sunken)
        self.line_3.setObjectName("line_3")
        self.verticalLayout.addWidget(self.line_3)
        self.state_name = QtWidgets.QLabel(self.centralwidget)
        sizePolicy = QtWidgets.QSizePolicy(QtWidgets.QSizePolicy.Preferred, QtWidgets.QSizePolicy.Fixed)
        sizePolicy.setHorizontalStretch(0)
        sizePolicy.setVerticalStretch(0)
        sizePolicy.setHeightForWidth(self.state_name.sizePolicy().hasHeightForWidth())
        self.state_name.setSizePolicy(sizePolicy)
        self.state_name.setMinimumSize(QtCore.QSize(100, 30))
        font = QtGui.QFont()
        font.setFamily("IBM Plex Sans")
        font.setPointSize(14)
        self.state_name.setFont(font)
        self.state_name.setText("")
        self.state_name.setAlignment(QtCore.Qt.AlignCenter)
        self.state_name.setObjectName("state_name")
        self.verticalLayout.addWidget(self.state_name)
        MainWindow.setCentralWidget(self.centralwidget)
        self.menubar = QtWidgets.QMenuBar(MainWindow)
        self.menubar.setGeometry(QtCore.QRect(0, 0, 692, 21))
        self.menubar.setObjectName("menubar")
        MainWindow.setMenuBar(self.menubar)
        self.statusbar = QtWidgets.QStatusBar(MainWindow)
        self.statusbar.setObjectName("statusbar")
        MainWindow.setStatusBar(self.statusbar)

        self.retranslateUi(MainWindow)
        QtCore.QMetaObject.connectSlotsByName(MainWindow)

    def retranslateUi(self, MainWindow):
        _translate = QtCore.QCoreApplication.translate
        MainWindow.setWindowTitle(_translate("MainWindow", "AREA - Install"))
        self.label.setText(_translate("MainWindow", "Установщик AREA - Student"))
        self.label_2.setText(_translate("MainWindow", "Перед установкой ознакомьтесь с политикой конфиденциальности"))
        self.privacy_policy_btn.setText(_translate("MainWindow", "Открыть политику конфиденциальности"))
        self.label_3.setText(_translate("MainWindow", "Путь установки"))
        self.change_path_btn.setText(_translate("MainWindow", "Выбрать путь"))
        self.label_4.setText(_translate("MainWindow", "Операционная система:"))
        self.system.setItemText(0, _translate("MainWindow", "Windows 32 bit"))
        self.system.setItemText(1, _translate("MainWindow", "Windows 64 bit"))
        self.install_btn.setText(_translate("MainWindow", "Установить"))
