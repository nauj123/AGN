# -*- coding: utf-8 -*-

# Form implementation generated from reading ui file 'PROYECTOAGN.ui'
#
# Created by: PyQt5 UI code generator 5.9.2
#
# WARNING! All changes made in this file will be lost!

from PyQt5 import QtCore, QtGui, QtWidgets

class Ui_MainWindow(object):
    def setupUi(self, MainWindow):
        MainWindow.setObjectName("MainWindow")
        MainWindow.resize(1024, 640)
        MainWindow.setMinimumSize(QtCore.QSize(1024, 640))
        MainWindow.setMaximumSize(QtCore.QSize(1024, 640))
        MainWindow.setStyleSheet("border-image: url(:/fondo/Imagenes/ventana_backgroudn_backgroudn.png);")
        self.centralwidget = QtWidgets.QWidget(MainWindow)
        self.centralwidget.setObjectName("centralwidget")
        self.boton4 = QtWidgets.QPushButton(self.centralwidget)
        self.boton4.setGeometry(QtCore.QRect(0, 380, 71, 61))
        self.boton4.setStyleSheet("border-image: url(:/btn4/Imagenes/metadatos-12.png);")
        self.boton4.setText("")
        self.boton4.setObjectName("boton4")
        self.boton5 = QtWidgets.QPushButton(self.centralwidget)
        self.boton5.setGeometry(QtCore.QRect(0, 500, 71, 61))
        self.boton5.setStyleSheet("border-image: url(:/btn5/Imagenes/database-12.png);")
        self.boton5.setText("")
        self.boton5.setObjectName("boton5")
        self.frame = QtWidgets.QFrame(self.centralwidget)
        self.frame.setGeometry(QtCore.QRect(90, 90, 921, 531))
        self.frame.setStyleSheet("border-image: url(:/Fondoes1/Imagenes/newinterface_backgroudn.png);")
        self.frame.setFrameShape(QtWidgets.QFrame.StyledPanel)
        self.frame.setFrameShadow(QtWidgets.QFrame.Raised)
        self.frame.setObjectName("frame")
        self.boton7 = QtWidgets.QPushButton(self.frame)
        self.boton7.setGeometry(QtCore.QRect(150, 140, 320, 80))
        font = QtGui.QFont()
        font.setFamily("Times New Roman")
        font.setPointSize(14)
        font.setBold(False)
        font.setWeight(50)
        self.boton7.setFont(font)
        self.boton7.setStyleSheet("border-image: url(:/btn7/Imagenes/Btn_Extraer_Individual_Pressed-02.png);")
        self.boton7.setObjectName("boton7")
        self.boton8 = QtWidgets.QPushButton(self.frame)
        self.boton8.setGeometry(QtCore.QRect(540, 140, 320, 80))
        font = QtGui.QFont()
        font.setFamily("Times New Roman")
        font.setPointSize(14)
        self.boton8.setFont(font)
        self.boton8.setStyleSheet("border-image: url(:/btn8/Imagenes/Btn_Extraer_Lote_Pressed-02.png);")
        self.boton8.setObjectName("boton8")
        self.progreso = QtWidgets.QProgressBar(self.frame)
        self.progreso.setGeometry(QtCore.QRect(190, 260, 691, 31))
        font = QtGui.QFont()
        font.setFamily("Times New Roman")
        font.setPointSize(18)
        self.progreso.setFont(font)
        self.progreso.setStyleSheet("")
        self.progreso.setProperty("value", 0)
        self.progreso.setObjectName("progreso")
        self.val1 = QtWidgets.QTextEdit(self.frame)
        self.val1.setGeometry(QtCore.QRect(420, 350, 71, 41))
        font = QtGui.QFont()
        font.setFamily("Times New Roman")
        font.setPointSize(20)
        self.val1.setFont(font)
        self.val1.setStyleSheet("")
        self.val1.setObjectName("val1")
        self.val2 = QtWidgets.QTextEdit(self.frame)
        self.val2.setGeometry(QtCore.QRect(560, 350, 61, 41))
        font = QtGui.QFont()
        font.setFamily("Times New Roman")
        font.setPointSize(20)
        self.val2.setFont(font)
        self.val2.setStyleSheet("")
        self.val2.setObjectName("val2")
        self.label = QtWidgets.QLabel(self.frame)
        self.label.setGeometry(QtCore.QRect(500, 360, 41, 21))
        font = QtGui.QFont()
        font.setFamily("Times New Roman")
        font.setPointSize(20)
        self.label.setFont(font)
        self.label.setObjectName("label")
        self.frame_2 = QtWidgets.QFrame(self.frame)
        self.frame_2.setGeometry(QtCore.QRect(0, 0, 921, 541))
        self.frame_2.setStyleSheet("background-image: url(:/Fondoes1/Imagenes/newinterface_backgroudn.png);")
        self.frame_2.setFrameShape(QtWidgets.QFrame.StyledPanel)
        self.frame_2.setFrameShadow(QtWidgets.QFrame.Raised)
        self.frame_2.setObjectName("frame_2")
        self.label_2 = QtWidgets.QLabel(self.frame_2)
        self.label_2.setGeometry(QtCore.QRect(330, 150, 301, 31))
        font = QtGui.QFont()
        font.setFamily("Times New Roman")
        font.setPointSize(20)
        self.label_2.setFont(font)
        self.label_2.setObjectName("label_2")
        self.metadatos1 = QtWidgets.QComboBox(self.frame_2)
        self.metadatos1.setGeometry(QtCore.QRect(50, 220, 811, 31))
        font = QtGui.QFont()
        font.setFamily("Times New Roman")
        font.setPointSize(15)
        self.metadatos1.setFont(font)
        self.metadatos1.setStyleSheet("border-image: url(:/fondoblanco/fondoblanco.png);")
        self.metadatos1.setObjectName("metadatos1")
        self.metadatos3 = QtWidgets.QComboBox(self.frame_2)
        self.metadatos3.setGeometry(QtCore.QRect(50, 350, 811, 31))
        font = QtGui.QFont()
        font.setFamily("Times New Roman")
        font.setPointSize(16)
        self.metadatos3.setFont(font)
        self.metadatos3.setStyleSheet("border-image: url(:/fondoblanco/fondoblanco.png);")
        self.metadatos3.setObjectName("metadatos3")
        self.label_5 = QtWidgets.QLabel(self.frame_2)
        self.label_5.setGeometry(QtCore.QRect(350, 290, 271, 21))
        font = QtGui.QFont()
        font.setFamily("Times New Roman")
        font.setPointSize(20)
        self.label_5.setFont(font)
        self.label_5.setObjectName("label_5")
        self.label_4 = QtWidgets.QLabel(self.frame_2)
        self.label_4.setGeometry(QtCore.QRect(270, 90, 481, 31))
        font = QtGui.QFont()
        font.setFamily("Times New Roman")
        font.setPointSize(21)
        self.label_4.setFont(font)
        self.label_4.setObjectName("label_4")
        self.frame_3 = QtWidgets.QFrame(self.frame_2)
        self.frame_3.setGeometry(QtCore.QRect(0, 0, 921, 541))
        self.frame_3.setStyleSheet("border-image: url(:/veninicio/Imagenes/vnew_backgroudn.png);")
        self.frame_3.setFrameShape(QtWidgets.QFrame.StyledPanel)
        self.frame_3.setFrameShadow(QtWidgets.QFrame.Raised)
        self.frame_3.setObjectName("frame_3")
        self.boton3 = QtWidgets.QPushButton(self.centralwidget)
        self.boton3.setGeometry(QtCore.QRect(0, 140, 71, 61))
        self.boton3.setStyleSheet("border-image: url(:/btn2/Imagenes/BtnExcel-12.png);")
        self.boton3.setText("")
        self.boton3.setObjectName("boton3")
        self.boton2 = QtWidgets.QPushButton(self.centralwidget)
        self.boton2.setGeometry(QtCore.QRect(0, 260, 71, 61))
        self.boton2.setStyleSheet("border-image: url(:/btn3/Imagenes/cargararchivoz-12.png);")
        self.boton2.setText("")
        self.boton2.setObjectName("boton2")
        MainWindow.setCentralWidget(self.centralwidget)

        self.retranslateUi(MainWindow)
        QtCore.QMetaObject.connectSlotsByName(MainWindow)

    def retranslateUi(self, MainWindow):
        _translate = QtCore.QCoreApplication.translate
        MainWindow.setWindowTitle(_translate("MainWindow", "SISTEMA DE GESTION DE METADATOS"))
        self.boton7.setWhatsThis(_translate("MainWindow", "<html><head/><body><p><span style=\" color:#ffffff;\">ARCHIVOS INDIVIDUALES</span></p></body></html>"))
        self.boton7.setText(_translate("MainWindow", "          ARCHIVOS INDIVIDUALES"))
        self.boton8.setText(_translate("MainWindow", "       ARCHIVOS POR LOTES"))
        self.progreso.setWhatsThis(_translate("MainWindow", "<html><head/><body><p align=\"center\"><br/></p></body></html>"))
        self.label.setText(_translate("MainWindow", "DE"))
        self.label_2.setText(_translate("MainWindow", "Seleccionar una Categoria"))
        self.label_5.setText(_translate("MainWindow", "Seleccionar Metadato"))
        self.label_4.setText(_translate("MainWindow", "BIENVENIDO A SU BASE DE DATOS"))

import fondo_rc
