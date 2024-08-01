"""
Description:
GUI основной страницы программы

Tasks:
Добавить создание текстового файла с логами (in progress);
Добавить выбор старых файлов (не скоро);
Сделать красивый интерфейс для программы;
Подумать над различными вариантами для разных лабораторий, тк приборы в основном закреплены за кабинетами,
чтобы не заморачиваться с инициализацией всех приборов сразу;
"""
import os
import win32com.client as win32
import pathlib
import time

from PyQt6.uic import loadUi
from PyQt6.QtWidgets import (QMainWindow, QDialog, QFileDialog)

from ChooseExcelDialog import ChooseExcelDialog
from Experiment import Experiment


class App(QMainWindow):
    """GUI основной страницы программы"""

    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)
        # Относительно долго грузится, можно разделить на 2 потока и добавить анимацию включения программы
        self.excel = win32.Dispatch('Excel.Application')  # Создаем COM объект
        self.excel.Visible = False  # Excel invisible

        loadUi('assets/mainIIC.ui', self)

        self.setWindowTitle('IIC Measuring Program')
        # self.setFixedSize(self.geometry().width(), self.geometry().height())

        self.ChooseButton.clicked.connect(self.onChooseExcelClicked)
        self.SettingsButton.clicked.connect(self.onSettingsClicked)
        self.CreateButton.clicked.connect(self.onCreateClicked)
        self.StartLineButton.clicked.connect(self.onStartLineClicked)
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
        self.data = {"TempName": "Нет шаблона",
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
        """Открывает настройки эксперимента"""
        pass

    def onCreateClicked(self):
        """Создаёт шаблон эксперимента"""
        self.ConsolePTE.appendPlainText(
            time.strftime("%H:%M:%S | ", time.localtime()) + 'НЕДОСТУПНО! Находится в разработке \n')
        pass

    def onStartLineClicked(self):
        """Задаёт начало строки, по дефолту выставляет начало строки на 11 (реализовать по созданию Excel)"""
        pass

    def onStartClicked(self):
        """Запускает эксперимент"""
        if self.working_flag:
            self.working_flag = False
            self.StartButton.setText('Старт')
        else:
            self.working_flag = True
            self.StartButton.setText('Стоп')

        # self.start_fuct()

    def onChooseExcelClicked(self):
        """Вызов диалога с созданием нового Excel"""
        self.ConsolePTE.appendPlainText(
            time.strftime("%H:%M:%S | ", time.localtime()) + 'Вызвали создание нового Excel')
        dlg = ChooseExcelDialog(self)
        dlg.exec()
