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

        # Путь к основной директории
        current_dir = os.path.dirname(os.path.abspath(__file__))
        # Загрузка ui, путем выхода в основную директорию

        loadUi(os.path.join(current_dir, '..', 'assets', 'chooseDialog.ui'), self)
        # loadUi('assets/chooseDialog.ui', self)

        self.app_instance = app_instance
        self.formatted_time = self.app_instance.formatted_time
        self.excel_settings = 'Smth'
        self.macros_status = None

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

    def log_message(self, message, exception=None):
        error_message = f"{self.formatted_time}{message}\n"
        if exception:
            error_message += f"{exception}\n"
        self.app_instance.ConsolePTE.appendPlainText(error_message)

    def open_macros_btn(self):
        """Метод для открытия макроса"""
        current_dir = os.path.dirname(os.path.abspath(__file__))
        dir_path = os.path.join(current_dir, '..', 'macroses')
        file_dir = QFileDialog.getOpenFileNames(self, 'Open File', dir_path, 'Macros (*txt)')
        if file_dir[0]:
            self.app_instance.data["MacrosName"] = file_dir[0][0]
            self.MacrosLabel.setText(file_dir[0][0])

    def open_temp_btn(self):
        """Метод для открытия шаблона"""
        current_dir = os.path.dirname(os.path.abspath(__file__))
        dir_path = os.path.join(current_dir, '..', 'templates')
        file_dir = QFileDialog.getOpenFileNames(self, 'Open File', dir_path, 'Excel file (*xltm)')
        if file_dir[0]:
            self.data["TempName"] = file_dir[0][0]
            self.TempLabel.setText(file_dir[0][0])

    def onAccepted(self):
        """Метод создающий новый Excel на основе заданных ранее параметров"""
        self.data["FileName"] = self.FileNameLE.text()
        # print(self.data["FileName"])

        script_path = os.path.dirname(os.path.abspath(__file__))
        file_path = os.path.join(script_path, '..', "Measurements", self.data["FileName"])  # without '.xlsm'

        if self.data["TempName"] == "Нет шаблона":
            temp_path = os.path.join(script_path, "..", "templates", "Base_temp.xltm")
            workbook = self.app_instance.excel.Workbooks.Open(temp_path)
        else:
            temp_path = os.path.join(script_path, "templates", self.data["TempName"])
            workbook = self.app_instance.excel.Workbooks.Open(self.data["TempName"])
            # workbook.SaveAs(file_path, 52)

        # Формат .xlsm будет при 52, а .xlsx при 51
        if self.data["MacrosName"] != "Нет макроса":
            try:
                vbacomponent = workbook.VBProject.VBComponents.Add(1)  # 1 = vbext_ct_StdModule
                vbacomponent.CodeModule.AddFromFile(self.data["MacrosName"])
                self.macros_status = f"с макросом \"{os.path.basename(self.data['MacrosName'])}\""

            except Exception as e:
                self.log_message('Макросы не добавлены, т.к.\n' +
                                 "в Центре управления безопасностью \n"
                                 "(Параметры макросов) необходимо поставить галочку на "
                                 "'Доверять доступ к объектной модели проектов VBA'", e)
                self.data["MacrosName"] = "Нет макроса"
                self.macros_status = "без макроса"
        else:
            self.macros_status = "без макроса"

        workbook.SaveAs(file_path, 52)
        self.app_instance.excel.Visible = True
        self.log_message(
            f"Создали файл \"{self.data['FileName']}\" по шаблону \"{os.path.basename(temp_path)}\" {self.macros_status}")

        # ! Добавить, если такой-то шаблон, то установки будут такие-то

    def onRejected(self):
        """Отмена создания Excel со сбросом заданных условий"""
        self.app_instance.data_reset_flag = not self.app_instance.data_reset_flag
        self.log_message('Отмена создания файла')

    def getExcelSettings(self):
        return self.excel_settings
