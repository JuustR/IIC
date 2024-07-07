"""
Диалоговое окно для создания нового файла Excel
"""
import os
import win32com.client as win32
import pathlib
import time

from PyQt6.uic import loadUi
from PyQt6.QtWidgets import (QDialog, QFileDialog)


class ChooseExcelDialog(QDialog):
    """Диалоговое окно для создания нового файла Excel"""

    def __init__(self, app_instance, parent=None):
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
        """Метод для открытия макроса"""
        dir_path = str(pathlib.Path.cwd())  # директория в которой находимся
        # dir_path = str(pathlib.Path.home()) # домашняя директоряи
        file_dir = QFileDialog.getOpenFileNames(self, 'Open File', dir_path + "\\macroses", 'Macros (*txt)')
        if file_dir[0]:
            self.app_instance.data["MacrosName"] = file_dir[0][0]
            self.MacrosLabel.setText(file_dir[0][0])

    def open_temp_btn(self):
        """Метод для открытия шаблона"""
        dir_path = str(pathlib.Path.cwd())  # директория в которой находимся
        # dir_path = str(pathlib.Path.home()) # домашняя директоряи
        file_dir = QFileDialog.getOpenFileNames(self, 'Open File', dir_path + "\\templates", 'Excel file (*xltm)')
        if file_dir[0]:
            self.data["TempName"] = file_dir[0][0]
            self.TempLabel.setText(file_dir[0][0])

    def onAccepted(self):
        """Метод создающий новый Excel на основе заданных ранее параметров"""
        self.data["FileName"] = self.FileNameLE.text()
        print(self.data["FileName"])

        script_path = os.path.abspath(__file__)
        file_path = os.path.dirname(script_path) + "\\Measurements\\" + self.data["FileName"]  # without '.xlsm'

        if self.data["TempName"] == "Нет шаблона":
            temp_path = os.path.dirname(script_path) + "\\templates\\Base_temp.xltm"
            workbook = self.app_instance.excel.Workbooks.Open(temp_path)
        else:
            workbook = self.app_instance.excel.Workbooks.Open(self.data["TempName"])
            temp_path = self.data["TempName"]
            # workbook.SaveAs(file_path, 52)

        # Формат .xlsm будет при 52, а .xlsx при 51
        if self.data["MacrosName"] != "Нет макроса":
            try:
                vbacomponent = workbook.VBProject.VBComponents.Add(1)  # 1 = vbext_ct_StdModule
                vbacomponent.CodeModule.AddFromFile(self.data["MacrosName"])
                print("Макросы успешно добавлены")  # Необязательно
            except:
                self.app_instance.ConsolePTE.appendPlainText(
                    time.strftime("%H:%M:%S | ", time.localtime()) + "Макросы не добавлены, т.к. ")
                self.app_instance.ConsolePTE.appendPlainText(
                    time.strftime("%H:%M:%S | ", time.localtime()) + """в Центре управления безопасностью 
(Параметры макросов) необходимо поставить галочку на Доверять доступ к объектной модели проектов VBA""")
                self.data["MacrosName"] = "Нет макроса"
        else:
            print("Макросы не добавлены")  # Необязательно

        workbook.SaveAs(file_path, 52)
        self.app_instance.excel.Visible = True

        if self.data["MacrosName"] == "Нет макроса":
            self.app_instance.ConsolePTE.appendPlainText(
                time.strftime("%H:%M:%S | ", time.localtime()) + "Создали файл Excel" +
                " по шаблону \"" + os.path.basename(temp_path) + "\" без макроса\n")
        else:
            self.app_instance.ConsolePTE.appendPlainText(
                time.strftime("%H:%M:%S | ", time.localtime()) + "Создали файл Excel" + " по шаблону \""
                + os.path.basename(temp_path) + "\" с макросом \"" + os.path.basename(self.data["MacrosName"]) + "\"\n")

    def onRejected(self):
        """Отмена создания Excel со сбросом заданных условий"""
        self.app_instance.data_reset_flag = not self.app_instance.data_reset_flag
        self.app_instance.ConsolePTE.appendPlainText(
            time.strftime("%H:%M:%S | ", time.localtime()) + 'Отмена создания файла')
