"""
Диалоговое окно для создания нового файла Excel
"""
import os
import time
import xlwings as xw
from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill, Side, Border
import os.path

from PyQt6.uic import loadUi
from PyQt6.QtWidgets import QDialog, QFileDialog, QStatusBar, QApplication, QMessageBox


class Create_Open_Excel(QDialog):
    """Диалоговое окно для создания нового файла Excel"""

    def __init__(self, app_instance, parent=None):
        super().__init__(parent)

        # Путь к основной директории
        current_dir = os.path.dirname(os.path.abspath(__file__))

        # Загрузка ui, путем выхода в основную директорию
        loadUi(os.path.join(current_dir, '..', "..", 'assets', 'chooseDialog2.ui'), self)

        self.app_instance = app_instance
        self.excel_settings = 'Smth'
        self.macros_status = None
        self.wb = None

        self.setWindowTitle('Choosing Template')

        self.statusbar = QStatusBar()
        self.MainVL.addWidget(self.statusbar)

        script_path = os.path.dirname(os.path.abspath(__file__))
        measurements_path = os.path.abspath(os.path.join(script_path, '..', "..", "Measurements"))
        self.label.setText(f"Файл: {measurements_path}\\")

        self.OpenMacrosBtn.clicked.connect(self.open_macros_btn)
        self.OpenTempBtn.clicked.connect(self.open_temp_btn)
        self.create_excel.clicked.connect(self.create_excel_func)
        self.open_excel.clicked.connect(self.open_excel_func)

        if self.app_instance.data_reset_flag:
            self.data = self.app_instance.base_data
            self.app_instance.data_reset_flag = not self.app_instance.data_reset_flag
        else:
            self.data = self.app_instance.data

        self.FileNameLE.setText(self.data["FileName"])
        self.MacrosLabel.setText(self.data["MacrosName"])
        self.TempLabel.setText(self.data["TempName"])

    def log_message(self, message, exception=None):
        error_message = time.strftime("%H:%M:%S | ", time.localtime()) + f"{message}\n"
        if exception:
            error_message += f"{exception}\n"
        self.app_instance.ConsolePTE.appendPlainText(error_message)

    def open_macros_btn(self):
        """
        Метод для открытия макроса
        """
        current_dir = os.path.dirname(os.path.abspath(__file__))
        dir_path = os.path.join(current_dir, '..', "..", 'macroses')
        file_dir = QFileDialog.getOpenFileNames(self, 'Open File', dir_path, 'Macros (*txt)')
        if file_dir[0]:
            self.app_instance.data["MacrosName"] = file_dir[0][0]
            self.MacrosLabel.setText(file_dir[0][0])

    def open_temp_btn(self):
        """
        Метод для открытия шаблона
        """
        current_dir = os.path.dirname(os.path.abspath(__file__))
        dir_path = os.path.join(current_dir, '..', "..", 'templates')
        file_dir = QFileDialog.getOpenFileNames(self, 'Open File', dir_path, 'Excel file (*xltm)')
        if file_dir[0]:
            self.data["TempName"] = file_dir[0][0]
            self.TempLabel.setText(file_dir[0][0])

    def getExcelSettings(self):
        return self.excel_settings

    def create_excel_func(self):

        try:
            self.data["FileName"] = self.FileNameLE.text()
            script_path = os.path.dirname(os.path.abspath(__file__))
            file_path = os.path.join(script_path, '..', "..", "Measurements", self.data["FileName"])  # without '.xlsm'

            # Проверка на наличие файлов с таким же именем
            if os.path.exists(file_path + ".xlsm"):
                self.log_message("Файл с таким именем уже существует, создайте другой")
                self.statusbar.showMessage("Файл с таким именем уже существует, создайте другой")
                return
            else:
                self.app_instance.wb = None

            if self.data["TempName"] == "Нет шаблона":
                temp_path = os.path.join(script_path, "..", "..", "templates", "Base_temp.xltm")
                with xw.App(visible=False) as app:
                    self.app_instance.wb = app.books.open(temp_path)
                    self.app_instance.wb.save(file_path + ".xlsm")
                    self.app_instance.wb.close()
                self.app_instance.wb = xw.Book(file_path + ".xlsm")
                self.app_instance.wb_path = file_path
            else:
                temp_path = os.path.join(script_path, "..", "..", "templates", self.data["TempName"])
                with xw.App(visible=False) as app:
                    self.app_instance.wb = app.books.open(temp_path)
                    self.app_instance.wb.save(file_path + ".xlsm")
                    self.app_instance.wb.close()
                self.app_instance.wb = xw.Book(file_path + ".xlsm")
                self.app_instance.wb_path = file_path

            if self.data["MacrosName"] != "Нет макроса":
                try:
                    vbacomponent = self.app_instance.wb.api.VBProject.VBComponents.Add(1)  # 1 = vbext_ct_StdModule
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

            self.app_instance.wb.save()
            self.log_message(
                f"Создали файл \"{self.data['FileName']}\" по шаблону \"{os.path.basename(temp_path)}\" {self.macros_status}")

            self.accept()
        except Exception as e:
            self.log_message(f'Неудалось создать эксель')


    def open_excel_func(self):
        try:
            file_name = str(self.FileNameLE.text()) + '.xlsm'
            script_path = os.path.dirname(os.path.abspath(__file__))
            file_path = os.path.abspath(os.path.join(script_path, '..', "..", "Measurements", file_name))
            self.app_instance.wb = xw.Book(file_path)
            self.app_instance.ws = self.app_instance.wb.sheets[0]

            self.accept()
        except Exception as e:
            self.statusbar.showMessage(f'Ошибка.\nНе удалось открыть Excel: {e}')