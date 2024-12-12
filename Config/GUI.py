"""

"""

import os
import win32com.client as win32
import pathlib
import time
import json

from PyQt6.uic import loadUi
from PyQt6.QtWidgets import (QMainWindow, QDialog, QFileDialog, QLineEdit)
from PyQt6.QtCore import QSettings

class App(QMainWindow):
    """GUI основной страницы программы"""

    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)
        # Относительно долго грузится, можно разделить на 2 потока и добавить анимацию включения программы
        self.excel = win32.Dispatch('Excel.Application')  # Создаем COM объект
        self.excel.Visible = False  # Excel invisible

        # Путь к основной директории
        current_dir = os.path.dirname(os.path.abspath(__file__))
        # Загрузка ui, путем выхода в основную директорию
        loadUi(os.path.join(current_dir, '..', 'assets', 'mainIIC2.ui'), self)

        self.setWindowTitle('IIC Measuring Program')

        # Хранилище настроек
        self.settings = QSettings("lab425", "IIC")

        # Подключаем основные кнопки к соответсвующим функциям
        self.ChooseButton.clicked.connect(self.onChooseExcelClicked)
        self.SettingsButton.clicked.connect(self.onSettingsClicked)
        self.CreateButton.clicked.connect(self.onCreateClicked)
        self.StartLineButton.clicked.connect(self.onStartLineClicked)
        self.StartButton.clicked.connect(self.onStartClicked)

        # Подключаем кнопки меню настроек
        self.PB_save_settings.clicked.connect(self.saveSettings)

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

        # Загрузка настроек для Seebeck+R
        self.load_tab1_settings()

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

    def saveSettings(self):
        """По кнопке сохраняет настройки программы для Seebeck+R"""
        self.settings.setValue("comboBox_scan", self.comboBox_scan.currentText())
        self.settings.setValue("comboBox_power", self.comboBox_power.currentText())
        self.settings.setValue("RB_newRead", self.RB_newRead.isChecked())
        self.settings.setValue("RB_oldRead", self.RB_oldRead.isChecked())
        self.settings.setValue("CB_average", self.CB_average.isChecked())
        self.settings.setValue("CB_rele", self.CB_rele.isChecked())

        for i in range(1, 7):  #Ch1 - Ch6
            self.settings.setValue(f"Ch{i}", self.findChild(QLineEdit, f"Ch{i}").text())
            self.settings.setValue(f"DelayCh{i}", self.findChild(QLineEdit, f"DelayCh{i}").text())

        self.settings.setValue("ChTerm", self.ChTerm.text())
        self.settings.setValue("ChTerm2", self.ChTerm2.text())
        self.settings.setValue("ChIP1", self.ChIP1.text())
        self.settings.setValue("ChIP2", self.ChIP2.text())
        self.settings.setValue("U_IP1", self.U_IP1.text())
        self.settings.setValue("U_IP2", self.U_IP2.text())
        self.settings.setValue("NPLC", self.NPLC.text())
        self.settings.setValue("TimeNPLC", self.TimeNPLC.text())
        self.settings.setValue("NRead", self.NRead.text())
        self.settings.setValue("Ncycles", self.Ncycles.text())
        self.settings.setValue("Nheat", self.Nheat.text())
        self.settings.setValue("Ncool", self.Ncool.text())
        self.settings.setValue("DelayS", self.DelayS.text())
        self.settings.setValue("DelayR", self.DelayR.text())
        self.settings.setValue("NR_Up", self.NR_Up.text())
        self.settings.setValue("NR_UpDown", self.NR_UpDown.text())

        #Добавление текста в консоль
        self.ConsolePTE.appendPlainText(
            time.strftime("%H:%M:%S | ", time.localtime()) + 'Сохранили настройки программы')

        #! Добавить сравнение изменений и вывод их в Excel, например
        #! 

    def load_tab1_settings(self):
        """Загружает сохранённые значения виджетов пока только для Seebeck+R"""
        self.comboBox_scan.setCurrentText(self.settings.value("comboBox_scan", ""))
        self.comboBox_power.setCurrentText(self.settings.value("comboBox_power", ""))
        self.RB_newRead.setChecked(self.settings.value("RB_newRead", "false") == "true")
        self.RB_oldRead.setChecked(self.settings.value("RB_oldRead", "false") == "true")
        self.CB_average.setChecked(self.settings.value("CB_average", "false") == "true")
        self.CB_rele.setChecked(self.settings.value("CB_rele", "false") == "true")

        for i in range(1, 7):
            self.findChild(QLineEdit, f"Ch{i}").setText(self.settings.value(f"Ch{i}", ""))
            self.findChild(QLineEdit, f"DelayCh{i}").setText(self.settings.value(f"DelayCh{i}", ""))

        self.ChTerm.setText(self.settings.value("ChTerm", ""))
        self.ChTerm2.setText(self.settings.value("ChTerm2", ""))
        self.ChIP1.setText(self.settings.value("ChIP1", ""))
        self.ChIP2.setText(self.settings.value("ChIP2", ""))
        self.U_IP1.setText(self.settings.value("U_IP1", ""))
        self.U_IP2.setText(self.settings.value("U_IP2", ""))
        self.NPLC.setText(self.settings.value("NPLC", ""))
        self.TimeNPLC.setText(self.settings.value("TimeNPLC", ""))
        self.NRead.setText(self.settings.value("NRead", ""))
        self.Ncycles.setText(self.settings.value("Ncycles", ""))
        self.Nheat.setText(self.settings.value("Nheat", ""))
        self.Ncool.setText(self.settings.value("Ncool", ""))
        self.DelayS.setText(self.settings.value("DelayS", ""))
        self.DelayR.setText(self.settings.value("DelayR", ""))
        self.NR_Up.setText(self.settings.value("NR_Up", ""))
        self.NR_UpDown.setText(self.settings.value("NR_UpDown", ""))