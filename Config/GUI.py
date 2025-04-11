"""
Окно основного графического интерфейса и его настройки

version 11.04.2025
© 04.2025 Frolov Dmitriy. All Rights Reserved. Any copying or distribution without the consent of the author is prohibited.
"""

import os
import win32com.client as win32
import time
import json

from PyQt6.uic import loadUi
from PyQt6.QtWidgets import QMainWindow, QLineEdit
from PyQt6.QtCore import QSettings, QTimer, pyqtSignal

from Config.ChooseExcelDialog import ChooseExcelDialog
from Config.Instruments import InstrumentConnection, ConnectionThread
from Config.Measurements import Measurements, MeasurementThread


class App(QMainWindow):
    """
    GUI основной страницы программы
    """
    log_signal = pyqtSignal(str)

    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)

        try:
            self.excel = win32.Dispatch('Excel.Application')
            self.excel.Visible = True  # Excel visible, тк invisible скрывает уже открытые файлы
        except Exception as e:
            print(f"Очистите gen_py\nМожно это сделать запустив программу gen_cache_clear.py\n{e}")

        # Путь к основной директории
        current_dir = os.path.dirname(os.path.abspath(__file__))
        # Загрузка ui, путем выхода в основную директорию
        loadUi(os.path.join(current_dir, '..', 'assets', 'mainIIC2.ui'), self)

        self.setWindowTitle('IIC Measuring Program')

        # Хранилище настроек
        self.settings = QSettings("lab425", "IIC")

        # Подключаем основные кнопки к соответсвующим функциям
        self.choose_button.clicked.connect(self.on_choose_excel_clicked)
        self.instruments_button.clicked.connect(self.on_instruments_clicked)
        self.create_button.clicked.connect(self.on_create_clicked)
        self.start_line_button.clicked.connect(self.on_start_line_clicked)
        self.start_button.clicked.connect(self.on_start_clicked)

        # checkBox'ы
        self.rele_cb.stateChanged.connect(self.rele_cb_clicked)
        self.instr1_cb.stateChanged.connect(self.instr1_cb_clicked)
        self.instr2_cb.stateChanged.connect(self.instr2_cb_clicked)

        # ComboBox'ы
        # self.combobox_scan.currentTextChanged.connect(self.combobox_scan_changed)
        # self.combobox_power.currentTextChanged.connect(self.combobox_power_changed)

        # Подключаем кнопки меню настроек
        self.save_settings_pb.clicked.connect(self.save_settings)

        # Начальные настройки консоли
        self.ConsolePTE.setPlainText(time.strftime("%d-%m-%Y %H:%M:%S", time.localtime()) + "\n" +
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
        self.data = {"TempName": "Нет шаблона",
                     "MacrosName": "Нет макроса",
                     "FileName": time.strftime("%d-%m-%Y_Example", time.localtime())}
        self.base_data = {"TempName": "Нет шаблона",
                          "MacrosName": "Нет макроса",
                          "FileName": time.strftime("%d-%m-%Y_Example", time.localtime())}

        # Флаги
        self.working_flag = False
        self.data_reset_flag = False
        self.settings_changed_flag = True
        self.startline_changed_flag = False
        self.excel_cash_flag = False

        self.wb_path = None
        self.wb = None
        self.ws = None
        self.inst_list = None
        self.powersource_list = None
        self.measurement = None
        self.settings_dict = {}
        self.start_time = time.time()
        self.excel_cash = []

        self.log_signal.connect(self.log_message)

        self.qtimer = QTimer(self)

        # Загрузка настроек для Seebeck+R
        self.load_tab1_settings()

        self.show()

    def log_message(self, message, exception=None):
        """
        Выводит в консоль pyqt логи

        :param message: Сообщение, которое будет выведено (str)
        :param exception: Вывод исключения/ошибки (exception)
        """
        error_message = time.strftime("%H:%M:%S | ", time.localtime()) + f"{message}\n"
        if exception:
            error_message += f"{exception}\n"
        self.ConsolePTE.appendPlainText(error_message)

    def instr1_cb_clicked(self):
        """
        Добавление ещё одного прибора
        """
        if self.instr1_cb.isChecked():
            self.ip_instr1.setEnabled(True)
            self.excel_instr1.setEnabled(True)
            self.instr1_connect.setEnabled(True)
        else:
            self.ip_instr1.setEnabled(False)
            self.excel_instr1.setEnabled(False)
            self.instr1_connect.setEnabled(False)

    def instr2_cb_clicked(self):
        """
        Добавление ещё одного прибора
        """
        if self.instr2_cb.isChecked():
            self.ip_instr2.setEnabled(True)
            self.excel_instr2.setEnabled(True)
            self.instr2_connect.setEnabled(True)
        else:
            self.ip_instr2.setEnabled(False)
            self.excel_instr2.setEnabled(False)
            self.instr2_connect.setEnabled(False)

    def rele_cb_clicked(self):
        """
        Подключение к релюшкам
        """
        if self.rele_cb.isChecked():
            self.n_r_up.setEnabled(False)
            self.n_r_updown.setEnabled(True)
        else:
            self.n_r_updown.setEnabled(False)
            self.n_r_up.setEnabled(True)

    def on_instruments_clicked(self):
        """
        Подключается ко всем доступным приборам, которые обнаружит
        """
        self.combobox_scan.clear()
        self.combobox_power.clear()
        ic = InstrumentConnection(self)

        self.connection_thread = ConnectionThread(ic)
        self.connection_thread.log_signal.connect(self.log_message)
        self.connection_thread.result_signal.connect(self.connection_finished)

        self.connection_thread.start()

    def connection_finished(self, inst_list, powersource_list):
        """
        Добавляет в соответствующие комбобоксы имена подключенных приборов

        :param inst_list: Список сканеров/мультиметров (list/dict)
        :param powersource_list: Список источников питания (list/dict)
        """
        self.inst_list = inst_list
        self.powersource_list = powersource_list

        connected_instruments = (', '.join(str(i) for i in self.inst_list) + ", " +
                                 ', '.join(str(i) for i in self.powersource_list))
        self.log_message('Подключенные приборы: ' + connected_instruments)

        for _ in self.inst_list:
            self.combobox_scan.addItem(_)
        for _ in self.powersource_list:
            self.combobox_power.addItem(_)

    def on_create_clicked(self):
        """
        Создаёт шаблон эксперимента
        """
        self.log_message('НЕДОСТУПНО! Находится в разработке')
        pass

    def on_start_line_clicked(self):
        """
        Задаёт начало строки, по дефолту выставляет начало строки на 11 (реализовать по созданию Excel)
        """
        self.startline_changed_flag = True

    def on_start_clicked(self):
        """
        Тут осуществляется запуск эксперимента
        """
        if self.working_flag:
            # Если поток активен, останавливаем его
            self.measurement_thread.stop()
            self.working_flag = False
            self.start_disable_le()
            self.start_button.setText('Старт')
            self.log_message("Измерения остановлены.")
        else:
            # Инициализация и запуск потока
            self.working_flag = True
            self.start_button.setText('Стоп')
            try:
                if self.wb is None:
                    self.log_message("Перед запуском убедитесь, что Excel создан")
                    self.working_flag = False
                    self.start_button.setText('Старт')
                    return
                else:
                    self.ws = self.wb.Worksheets(1)
                if not self.inst_list or not self.powersource_list:
                    self.log_message("Перед запуском убедитесь, что приборы и источники питания подключены")
                    self.working_flag = False
                    self.start_button.setText('Старт')
                    return
                self.measurement = Measurements(self)
                self.measurement_thread = MeasurementThread(self.measurement)
                self.measurement.update_excel_signal.connect(self.update_excel)
                self.measurement.update_values_signal.connect(self.update_values)
                self.measurement_thread.log_signal.connect(self.log_message)
                self.measurement_thread.finished_signal.connect(self.measurement_finished)
                self.measurement_thread.start()
                self.start_disable_le()
                self.log_message("Начало измерений")
            except Exception as e:
                self.log_message("Ошибка запуска измерений.", e)
                self.working_flag = False
                self.start_disable_le()
                self.start_button.setText('Старт')

    def update_values(self, dict):
        """
        Обновляет различные параметры, которые получает из Measurements

        :param dict: Словарь в котором содержатся пункты, которые нужно обновить
        """
        if "start_line_le" in dict:
            self.start_line_le.setText(dict["start_line_le"])

    def update_excel(self, row, col, value):
        """
        Выполняет изменения в Excel в основном потоке

        :param row: Номер строки (int)
        :param col: Номер или имя столбца (int/str)
        :param value: Значение, которое будет записано в ячейку (object)
        """
        try:
            self.ws.Cells(row, col).Value = value
            if self.excel_cash_flag:
                while self.excel_cash:
                    x = self.excel_cash[0]
                    try:
                        self.ws.Cells(x[0], x[1]).Value = x[2]
                        self.excel_cash.pop(0)
                    except Exception as e:
                        self.excel_cash_flag = True
                        print(e)
                        break
        except Exception as e:
            self.excel_cash.append([row, col, value])
            self.excel_cash_flag = True
            # self.log_message(f"Ошибка записи в Excel: {e}")
            print(f"Ошибка записи в Excel: {e}")

    def measurement_finished(self):
        self.measurement.pause()
        self.log_message("Измерение остановлено")
        self.working_flag = False
        self.start_button.setText('Старт')

    def pause(self):
        if self.measurement_thread and self.measurement_thread.isRunning():
            self.measurement_thread.stop()
            self.measurement_thread.wait()
        self.measurement.pause()
        self.working_flag = False
        self.start_button.setText('Старт')

    def on_choose_excel_clicked(self):
        """
        Вызов диалога с созданием нового Excel
        """
        self.log_message('Вызвали создание нового Excel')
        dlg = ChooseExcelDialog(self)
        dlg.exec()


    def save_settings(self):
        """
        По кнопке сохраняет настройки программы для Seebeck+R
        """
        self.settings.setValue("combobox_scan", self.combobox_scan.currentText())
        self.settings.setValue("combobox_power", self.combobox_power.currentText())
        self.settings.setValue("rele_cb", self.rele_cb.isChecked())

        for i in range(1, 7):  # Ch1 - Ch6
            self.settings.setValue(f"ch{i}", self.findChild(QLineEdit, f"ch{i}").text())
            self.settings.setValue(f"delay_ch{i}", self.findChild(QLineEdit, f"delay_ch{i}").text())
            self.settings.setValue(f"range_ch{i}", self.findChild(QLineEdit, f"range_ch{i}").text())
            self.settings.setValue(f"nplc_ch{i}", self.findChild(QLineEdit, f"nplc_ch{i}").text())

        self.settings.setValue("n_read_ch12", self.n_read_ch12.text())
        self.settings.setValue("n_read_ch34", self.n_read_ch34.text())
        self.settings.setValue("n_read_ch56", self.n_read_ch56.text())
        self.settings.setValue("ch_term1", self.ch_term1.text())
        # self.settings.setValue("ch_term2", self.ch_term2.text())
        self.settings.setValue("delay_term", self.delay_term.text())
        self.settings.setValue("range_term", self.range_term.text())
        self.settings.setValue("nplc_term", self.nplc_term.text())
        self.settings.setValue("ch_ip1", self.ch_ip1.text())
        self.settings.setValue("ch_ip2", self.ch_ip2.text())
        self.settings.setValue("u_ip1", self.u_ip1.text())
        self.settings.setValue("u_ip2", self.u_ip2.text())
        self.settings.setValue("n_cycles", self.n_cycles.text())
        self.settings.setValue("n_heat", self.n_heat.text())
        self.settings.setValue("n_cool", self.n_cool.text())
        self.settings.setValue("pause_s", self.pause_s.text())
        self.settings.setValue("pause_r", self.pause_r.text())
        self.settings.setValue("n_r_up", self.n_r_up.text())
        self.settings.setValue("n_r_updown", self.n_r_updown.text())
        self.settings.setValue("ip_rigol", self.ip_rigol.text())
        self.settings.setValue("r_cell", self.r_cell.text())

        # Добавление текста в консоль
        self.log_message('Сохранили настройки программы')
        if self.working_flag:
            self.log_message('Настройки будут применены в следующем цикле')

        self.settings_changed_flag = True
        # ! Добавить сравнение изменений и вывод их в Excel, например
        # !

    def load_tab1_settings(self):
        """
        Загружает сохранённые значения виджетов пока только для Seebeck+R
        """
        self.combobox_scan.setCurrentText(self.settings.value("combobox_scan", ""))
        self.combobox_power.setCurrentText(self.settings.value("combobox_power", ""))
        self.rele_cb.setChecked(self.settings.value("rele_cb", "false") == "true")

        for i in range(1, 7):
            self.findChild(QLineEdit, f"ch{i}").setText(self.settings.value(f"ch{i}", ""))
            self.findChild(QLineEdit, f"delay_ch{i}").setText(self.settings.value(f"delay_ch{i}", ""))
            self.findChild(QLineEdit, f"range_ch{i}").setText(self.settings.value(f"range_ch{i}", ""))
            self.findChild(QLineEdit, f"nplc_ch{i}").setText(self.settings.value(f"nplc_ch{i}", ""))

        self.n_read_ch12.setText(self.settings.value("n_read_ch12", ""))
        self.n_read_ch34.setText(self.settings.value("n_read_ch34", ""))
        self.n_read_ch56.setText(self.settings.value("n_read_ch56", ""))
        self.ch_term1.setText(self.settings.value("ch_term1", ""))
        # self.ch_term2.setText(self.settings.value("ch_term2", ""))
        self.ch_ip1.setText(self.settings.value("ch_ip1", ""))
        self.ch_ip2.setText(self.settings.value("ch_ip2", ""))
        self.u_ip1.setText(self.settings.value("u_ip1", ""))
        self.u_ip2.setText(self.settings.value("u_ip2", ""))
        self.delay_term.setText(self.settings.value("delay_term", ""))
        self.range_term.setText(self.settings.value("range_term", ""))
        self.nplc_term.setText(self.settings.value("nplc_term", ""))
        self.n_cycles.setText(self.settings.value("n_cycles", ""))
        self.n_heat.setText(self.settings.value("n_heat", ""))
        self.n_cool.setText(self.settings.value("n_cool", ""))
        self.pause_s.setText(self.settings.value("pause_s", ""))
        self.pause_r.setText(self.settings.value("pause_r", ""))
        self.n_r_up.setText(self.settings.value("n_r_up", ""))
        self.n_r_up.setText(self.settings.value("n_r_up", ""))
        self.ip_rigol.setText(self.settings.value("ip_rigol", ""))
        self.r_cell.setText(self.settings.value("r_cell", ""))

        # (мб необ) Копирование настроек в словарь
        self.copy_settings_to_dict()

    def copy_settings_to_dict(self):
        """
        Копирует настройки из self.settings в словарь settings_dict
        """

        def copy_to_dict(element):
            self.settings_dict[element] = self.settings.value(element)

        keys = [
            "combobox_scan", "combobox_power", "rele_cb",
            "n_read_ch12", "n_read_ch34", "n_read_ch56",
            "ch_term1", "ch_term2", "delay_term", "range_term", "nplc_term",
            "ch_ip1", "ch_ip2", "u_ip1", "u_ip2", "n_cycles",
            "n_heat", "n_cool", "pause_s", "pause_r",
            "n_r_up", "n_r_updown", "ip_rigol"
        ]

        for i in range(1, 7):
            keys.extend([f"ch{i}", f"delay_ch{i}", f"range_ch{i}", f"nplc_ch{i}"])

        for key in keys:
            copy_to_dict(key)

    def start_disable_le(self):
        """
        Отключение фрагментов интерфейса при старте измерений
        """
        widgets_to_disable = [self.combobox_scan, self.combobox_power, self.ip_rigol, self.n_read_ch12, self.n_read_ch34,
                              self.n_read_ch56, self.n_cycles, self.n_heat, self.n_cool, self.ch1, self.ch2, self.ch3,
                              self.ch4, self.ch5, self.ch6, self.ch_term1, self.instruments_button, self.choose_button,
                              self.create_button]
        if self.working_flag:
            for widget in widgets_to_disable:
                widget.setEnabled(False)
            if self.rele_cb.isChecked():
                self.n_r_updown.setEnabled(False)
            else:
                self.n_r_up.setEnabled(False)
        else:
            for widget in widgets_to_disable:
                widget.setEnabled(True)
            if self.rele_cb.isChecked():
                self.n_r_updown.setEnabled(True)
            else:
                self.n_r_up.setEnabled(True)


    def closeEvent(self, event):
        """
        Завершает работу Excel перед выходом
        """
        try:
            if self.wb:
                self.wb.SaveAs(self.wb_path, 52)
                self.wb.Close(SaveChanges=1)
            if self.excel:
                self.excel.Quit()
            print("Excel успешно закрыт")
        except Exception as e:
            print(f"Ошибка при закрытии Excel: {e}")
        finally:
            event.accept()