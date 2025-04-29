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

    def onAccepted(self):
        """
        Метод создающий новый Excel на основе заданных ранее параметров
        """
        self.data["FileName"] = self.FileNameLE.text()
        script_path = os.path.dirname(os.path.abspath(__file__))
        file_path = os.path.join(script_path, '..', "..", "Measurements", self.data["FileName"])  # without '.xlsm'

        # Проверка на наличие файлов с таким же именем
        if os.path.exists(file_path + ".xlsm"):
            self.log_message("Файл с таким именем уже существует, создайте другой")
            return
        else:
            self.app_instance.wb = None

        if self.data["TempName"] == "Нет шаблона":
            temp_path = os.path.join(script_path, "..", "..", "templates", "Base_temp.xltm")
            # workbook = self.excel.Workbooks.Open(temp_path)
            self.app_instance.wb = self.app_instance.excel.Workbooks.Open(temp_path)
            self.app_instance.wb_path = file_path
        else:
            temp_path = os.path.join(script_path, "templates", self.data["TempName"])
            # workbook = self.excel.Workbooks.Open(self.data["TempName"])
            self.app_instance.wb = self.app_instance.excel.Workbooks.Open(self.data["TempName"])
            self.app_instance.wb_path = file_path

        # Формат .xlsm будет при 52, а .xlsx при 51
        if self.data["MacrosName"] != "Нет макроса":
            try:
                vbacomponent = self.wb.VBProject.VBComponents.Add(1)  # 1 = vbext_ct_StdModule
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

        self.app_instance.wb.SaveAs(file_path, 52)
        self.app_instance.excel.Visible = True
        self.log_message(
            f"Создали файл \"{self.data['FileName']}\" по шаблону \"{os.path.basename(temp_path)}\" {self.macros_status}")

        # ! Добавить, если такой-то шаблон, то установки будут такие-то


    def onRejected(self):
        """
        Отмена создания Excel со сбросом заданных условий
        """
        self.app_instance.data_reset_flag = not self.app_instance.data_reset_flag
        self.log_message('Отмена создания файла')

    def getExcelSettings(self):
        return self.excel_settings

    def create_excel_func(self):
        try:

            file_name = str(self.FileNameLE.text()) + '.xlsm'

            book = Workbook()
            sheet = book.active

            thin_border = Border(left=Side(style='thin'),
                                 right=Side(style='thin'),
                                 top=Side(style='thin'),
                                 bottom=Side(style='thin'))

            font = Font(b=True, size=12, color="000000")
            fill = PatternFill("solid", fgColor="FFFF00")

            sheet['A1'] = 'Время'
            sheet['B1'] = 'V1'
            sheet['C1'] = 'V2'
            sheet['D1'] = 'V3'
            sheet['E1'] = 'V4'
            sheet['F1'] = 'Системное время'
            sheet['G1'] = 'Комментарии Python'
            sheet['H1'] = 'Для комментариев'

            # sheet.column_dimensions['D'].width = 20
            sheet.column_dimensions['H'].width = 25
            sheet.column_dimensions['F'].width = 25
            sheet.column_dimensions['L'].width = 20
            sheet.column_dimensions['M'].width = 25
            sheet.column_dimensions['N'].width = 25

            for i in range(1, 25):
                sheet.cell(row=1, column=i).font = font
                sheet.cell(row=1, column=i).fill = fill
                sheet.cell(row=1, column=i).alignment = Alignment(horizontal='center')
                sheet.cell(row=1, column=i).border = thin_border

            if os.path.isfile("C:\\Python\\" + file_name):
                error = QMessageBox()
                error.setWindowTitle('Ошибка')
                error.setText('Такой файл уже существует, выберите другое имя')
                # error.setIcon(QMessageBox.Warning)
                # error.setStandardButtons(QMessageBox.Ok)
                error.exec()
            else:
                book.save("C:\\Python\\" + file_name)
                # os.startfile("C:\\Python\\" + file_name)
                book.close()
                self.time_open = time.time()
                while time.time() - self.time_open < 2:
                    pass
                self.wb = xw.Book(f"C:\\Python\\{file_name}")
                self.ws = self.wb.sheets[0]

        except Exception as e:
            QMessageBox.critical(self, 'Ошибка', f'Не удалось создать файл: {e}')

    def open_excel_func(self):
        try:
            file_name = str(self.FileNameLE.text()) + '.xlsm'
            self.wb = xw.Book(f"C:\\IIC\\Measurements\\{file_name}")
            self.ws = self.wb.sheets[0]
            # xw.App(visible=True)
        except Exception as e:
            self.statusbar.showMessage(f'Ошибка.\nНе удалось открыть Excel: {e}')