"""
Добавить создание текстового файла с логами

"""
import pathlib
import time
from typing import Union

from PyQt6 import uic, QtCore, QtTest, QtGui, QtWidgets
from PyQt6.uic import loadUi
from PyQt6.QtWidgets import (QApplication, QWidget, QMessageBox, QMainWindow,
                             QDialog, QDialogButtonBox, QLabel, QFileDialog, QVBoxLayout, QLineEdit, QPushButton,
                             QHBoxLayout, QScrollArea, QFormLayout, QGroupBox, QComboBox)
from PyQt6.QtCore import QFile, QIODevice, Qt
from PyQt6.QtGui import QIcon, QPixmap
from PyQt6.QtGui import QIcon, QMovie


class App(QMainWindow):

    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)
        uic.loadUi('assets/mainIIC.ui', self)
        self.init()

    def init(self):
        self.setWindowTitle('IIC Measuring Program')
        # self.setFixedSize(self.geometry().width(), self.geometry().height())
        # self.SettingsButton.clicked.connect(self.settings_button)
        # self.InstButton.clicked.connect(self.inst_settings)
        # self.NoClickButton.clicked.connect(self.no_click)
        # self.StartStopButton.clicked.connect(self.start_stop)
        # self.ApplaySettings.clicked.connect(self.applay_settings)
        # self.OpenFile.clicked.connect(self.open_file)
        # self.CreateFile.clicked.connect(self.create_file)
        # self.timer_start = QtCore.QTimer()
        # self.timer_start.timeout.connect(lambda: self.logic.measurement())

        self.ChooseButton.clicked.connect(self.onChooseExcelClicked)
        # self.SettingsButton.clicked.connect(self.onSettingsClicked)
        # self.CreateButton.clicked.connect(self.onCreateClicked)
        # self.StartLineButton.clicked.connect(self.onStartLineClicked)
        self.StartButton.clicked.connect(self.onStartClicked)

        # Time and console start settings
        formatted_time = time.strftime("%d-%m-%Y %H:%M:%S", time.localtime())
        self.ConsolePTE.setPlainText(formatted_time + "\n" + """
Здравствуйте! Для начала работы:
1. Выберите/создайте шаблон Excel
2. Настройте приборы и Excel
  2.1. Выберите каналы
  2.2. Настройте параметры
  2.3. Подключитель к приборам
3.(необ.) Задайте начальную строчку Excel
4. Запускайте измерения
        """)

        # Flag for start button
        self.working_flag = False

        # self.movie = QMovie("C:/Users/Comp.ROMAN/Desktop/giphy.gif")
        # self.label.setMovie(self.movie)
        # self.movie.start()

        self.show()

    def onSettingsClicked(self):
        pass

    def onCreateClicked(self):
        pass

    def onStartLineClicked(self):
        # Задаёт начало строки, по дефолту выставляет начало строки на 11 (реализовать по созданию Excel)
        pass

    def onStartClicked(self):
        if self.working_flag:
            self.working_flag = False
            self.StartButton.setText('Старт')
        else:
            self.working_flag = True
            self.StartButton.setText('Стоп')

        #self.start_fuct()

    def onChooseExcelClicked(self):
        self.ConsolePTE.appendPlainText(time.strftime("%H:%M:%S | ", time.localtime()) + 'Вызвали выбор шаблона Excel \n')
        dlg = ChooseExcelDialog(self)
        dlg.exec()


class ChooseExcelDialog(QDialog):

    def __init__(self, parent=None):
        super().__init__(parent)
        loadUi('assets/chooseDialog.ui', self)
