"""
Добавить создание текстового файла с логами

"""
import os
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
    """GUI основной страницы программы"""
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
        self.ConsolePTE.setPlainText(formatted_time + "\n" +
                                     """
Здравствуйте! Для начала работы:
1. Выберите/создайте шаблон Excel
2. Настройте приборы и Excel
  2.1. Выберите каналы
  2.2. Настройте параметры
  2.3. Подключитель к приборам
3.(необ.) Задайте начальную строчку Excel
4. Запускайте измерения
        """)

        # Dictionary for requests to other classes
        # data - changeable, base_data - unchangeable REMEMBER IT!!!
        # Вообще выглядит херово, я бы как-то изменил
        # In use: FileName, TempName, MacrosName
        self.data = {"TempName": "Нет шаблона1",
                     "MacrosName": "Нет макроса",
                     "FileName": "Example"}
        self.base_data = {"TempName": "Нет шаблона",
                     "MacrosName": "Нет макроса",
                     "FileName": "Example"}

        # Flag for start button
        self.working_flag = False
        self.data_reset_flag = False

        self.show()

    def onSettingsClicked(self):
        pass

    def onCreateClicked(self):
        self.ConsolePTE.appendPlainText(
            time.strftime("%H:%M:%S | ", time.localtime()) + 'НЕДОСТУПНО! Находится в разработке \n')
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

        # self.start_fuct()

    def onChooseExcelClicked(self):
        self.ConsolePTE.appendPlainText(
            time.strftime("%H:%M:%S | ", time.localtime()) + 'Вызвали выбор шаблона Excel')
        dlg = ChooseExcelDialog(self)
        dlg.exec()


class ChooseExcelDialog(QDialog):

    def __init__(self, app_instance,  parent=None):
        super().__init__(parent)
        loadUi('assets/chooseDialog.ui', self)

        self.app_instance = app_instance

        self.setWindowTitle('Choosing Template')
        self.OpenMacrosBtn.clicked.connect(self.open_macros_btn)
        self.OpenTempBtn.clicked.connect(self.open_temp_btn)
        self.buttonBox.accepted.connect(self.onAccepted)
        self.buttonBox.rejected.connect(self.onRejected)

        if self.app_instance.data_reset_flag:
            self.data = self.app_instance.base_data
            self.app_instance.data_reset_flag = not self.app_instance.data_reset_flag
        else:
            self.data = self.app_instance.data

        self.FileNameLE.setText(self.data["FileName"])
        self.MacrosLabel.setText(self.data["MacrosName"])
        self.TempLabel.setText(self.data["TempName"])


    def open_macros_btn(self):
        dir_path = str(pathlib.Path.cwd())  # директория в которой находимся
        # dir_path = str(pathlib.Path.home()) # домашняя директоряи
        file_dir = QFileDialog.getOpenFileNames(self, 'Open File', dir_path + "\\macroses", 'Macros (*txt)')
        if file_dir[0]:
            self.app_instance.data["MacrosName"] = file_dir[0][0]
            self.MacrosLabel.setText(file_dir[0][0])

    def open_temp_btn(self):
        dir_path = str(pathlib.Path.cwd())  # директория в которой находимся
        # dir_path = str(pathlib.Path.home()) # домашняя директоряи
        file_dir = QFileDialog.getOpenFileNames(self, 'Open File', dir_path + "\\templates", 'Excel file (*xltm)')
        if file_dir[0]:
            self.data["TempName"] = file_dir[0][0]
            self.TempLabel.setText(file_dir[0][0])

    def onAccepted(self):
        self.data["FileName"] = self.FileNameLE.text()
        #
        # script_path = os.path.abspath(__file__)
        # file_path = os.path.dirname(script_path) + "\\Measurements\\" + self.data["FileName"]  # without '.xlsm'
        #
        # if self.data["TempName"] == None:
        #     workbook = self.excel.Workbooks.Open(os.path.dirname(script_path) + "\\templates\\Base_temp.xltm")
        # else:
        #     workbook = self.excel.Workbooks.Open(self.data[4])
        #     # workbook.SaveAs(file_path, 52)
        #
        # # Формат .xlsm будет при 52, а .xlsx при 51
        # if self.data[2]:
        #     try:
        #         vbacomponent = workbook.VBProject.VBComponents.Add(1)  # 1 = vbext_ct_StdModule
        #         vbacomponent.CodeModule.AddFromFile("C:\\Dima\\INH\\VBAcode.txt")
        #         print("Макросы успешно добавлены")
        #     except:
        #         print("Макросы не добавлены")
        #         print(
        #             "В Центре управления безопасностью поставить галочку на Доверять доступ к объектной модели проектов VBA")
        # else:
        #     print("Макросы не добавлены")
        # workbook.SaveAs(file_path, 52)

        self.app_instance.ConsolePTE.appendPlainText(
            time.strftime("%H:%M:%S | ", time.localtime()) + 'Создали файл Excel')

    def onRejected(self):
        self.app_instance.data_reset_flag = not self.app_instance.data_reset_flag
        self.app_instance.ConsolePTE.appendPlainText(
            time.strftime("%H:%M:%S | ", time.localtime()) + 'Отмена')

