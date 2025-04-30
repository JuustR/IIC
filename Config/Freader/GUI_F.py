"""
Новое GUI, основанное на Freader

Ver 0.01 from 29.04.2025

Задачи:

1)
"""

import sys
import time
import os
import xlwings as xw
import pyvisa
from PyQt6.QtWidgets import QRadioButton
from PyQt6.uic import loadUi
from PyQt6.QtWidgets import QApplication, QMainWindow, QMessageBox, QPushButton, QCheckBox, QLineEdit
from PyQt6.QtCore import pyqtSignal, QThread, QSettings, QTimer


from Config.Freader.Meas_F import ExcelWriterThread
from Config.Freader.Instruments_F import InstrumentConnection, ConnectionThread
from Config.Freader.Create_Excel import Create_Open_Excel


class App(QMainWindow):
    """
    Documentation
    """
    log_signal = pyqtSignal(str)

    def __init__(self):
        super().__init__()

        self.rm = pyvisa.ResourceManager()

        # Путь к основной директории
        current_dir = os.path.dirname(os.path.abspath(__file__))
        # Загрузка ui, путем выхода в основную директорию
        loadUi(os.path.join(current_dir, '..', '..', 'assets', 'mainIIC3.ui'), self)

        # loadUi('daq_reader_by_df.ui', self)  # Загружаем UI напрямую в self
        self.setWindowTitle("Seebeck + R Measurement Programm by DIF")

        self.settings = QSettings("lab425", "SRMP_by_DIF")

        # self.excel_name.setText(time.strftime("%d%m%Y ", time.localtime()) + "_Example")

        # Анимация
        self.animation_timer = QTimer()
        self.animation_timer.setInterval(500)
        self.animation_timer.timeout.connect(self.animate_text)
        self.animation_num = 2

        self.wb_path = None
        self.wb = None
        self.ws = None
        self.inst_list = None
        self.powersource_list = None
        self.measurement = None
        self.settings_dict = {}
        self.start_time = time.time()
        self.cache_values = []  # self.excel_cash

        self.log_signal.connect(self.log_message)

        self.working_flag = False
        self.data_reset_flag = False
        self.settings_changed_flag = True
        self.startline_changed_flag = False
        self.excel_cash_flag = False

        # !Было
        # self.wb = None
        # self.ws = None
        # self.daq = None
        # self.time_open = None  # Время создание файла
        # self.start_time = time.time()
        # self.usb_resources = None
        # self.cache_values = []

        # self.start_time_flag = False
        # self.ip_daq_flag = True
        # self.read_one_flag = False
        # self.rts_flag = False  # Ready to Start
        # self.start_flag = False

        # Подключаем основные кнопки к соответсвующим функциям
        self.choose_button.clicked.connect(self.on_choose_excel_clicked)
        self.instruments_button.clicked.connect(self.on_instruments_clicked)
        # self.create_button.clicked.connect(self.on_create_clicked)
        # self.start_line_button.clicked.connect(self.on_start_line_clicked)
        self.start_button.clicked.connect(self.start)
        self.pause_button.clicked.connect(self.pause)
        # checkBox'ы
        self.rele_cb.stateChanged.connect(self.rele_cb_clicked)
        self.instr1_cb.stateChanged.connect(self.instr1_cb_clicked)
        self.instr2_cb.stateChanged.connect(self.instr2_cb_clicked)

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

        # !Переписать
        # data - changeable, base_data - unchangeable REMEMBER IT!!!
        self.data = {"TempName": "Нет шаблона",
                     "MacrosName": "Нет макроса",
                     "FileName": time.strftime("%d-%m-%Y_Example", time.localtime())}
        self.base_data = {"TempName": "Нет шаблона",
                          "MacrosName": "Нет макроса",
                          "FileName": time.strftime("%d-%m-%Y_Example", time.localtime())}

        # !Было
        # self.create_excel.clicked.connect(self.create_excel_func)
        # self.open_excel.clicked.connect(self.open_excel_func)
        # self.start_button.clicked.connect(self.start)

        # self.connect_button.clicked.connect(self.connect)
        # self.read_one.clicked.connect(self.read_one_func)
        # self.save_button.clicked.connect(self.save_settings)

        self.mainthread = None

        self.load_tab1_settings()

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
        # !Добавить проверку в измерения, чтоб отключать взаимодействие с релюшками
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

    def on_choose_excel_clicked(self):
        """
        Вызов диалога с созданием нового Excel
        """
        self.log_message('Вызвали создание нового Excel')
        dlg = Create_Open_Excel(self)
        dlg.exec()

    def start(self):
        self.start_button.setText("Измеряется")
        self.start_flag = True
        # self.disable_enable_ch()
        self.start_disable_le()
        self.animation_timer.start()
        try:
            if self.wb is None:
                self.log_message("Перед запуском убедитесь, что Excel создан")
                self.statusbar.showMessage("Перед запуском убедитесь, что Excel создан")
                self.pause()
                return
            if self.inst_list:
                if self.mainthread is None or not self.mainthread.isRunning():
                    self.mainthread = ExcelWriterThread(self)

                    # self.mainthread.update_values_signal.connect(self.update_values)
                    self.mainthread.update_excel_signal.connect(self.update_excel)
                    self.mainthread.stop_signal.connect(self.pause)
                    self.mainthread.start()
                else:
                    self.pause()
            else:
                print("Poka net")
        except Exception as e:
            print(e)
            self.pause()

    def update_excel(self, row, col, value):
        try:
            for _ in self.cache_values:
                self.ws.cells(_[0], _[1]).value = _[2]
            self.cache_values.clear()

            self.ws.cells(row, col).value = value

        except:
            # [row, col]
            self.cache_values.append([row, col, value])
            print(f"Cache: {self.cache_values}")

    def update_values(self, num):
        self.start_line_le.setText(f'{num}')

    def pause(self):
        self.start_flag = False
        # self.disable_enable_ch()
        self.start_button.setEnabled(True)

        if self.animation_timer.isActive():
            self.animation_timer.stop()
            self.start_button.setText("Старт")
            self.animation_num = 2

        if self.mainthread and self.mainthread.isRunning():
            self.mainthread.stop()
            self.mainthread.wait()
            self.mainthread = None

        self.statusbar.showMessage("Измерение остановлено")

    def disable_enable_ch(self):
        if self.start_flag:
            for i in range(1, 21):
                current_ch = self.findChild(QCheckBox, f"ch{i}")
                current_dcv = self.findChild(QRadioButton, f"dcv{i}")
                current_fres = self.findChild(QRadioButton, f"fres{i}")
                current_ch.setEnabled(False)
                current_dcv.setEnabled(False)
                current_fres.setEnabled(False)
        else:
            for i in range(1, 21):
                current_ch = self.findChild(QCheckBox, f"ch{i}")
                current_dcv = self.findChild(QRadioButton, f"dcv{i}")
                current_fres = self.findChild(QRadioButton, f"fres{i}")
                current_ch.setEnabled(True)
                current_dcv.setEnabled(True)
                current_fres.setEnabled(True)

    def animate_text(self):
        if self.animation_num == 1:
            self.start_button.setText("Измеряется")
        elif self.animation_num == 2:
            self.start_button.setText("Измеряется.")
        elif self.animation_num == 3:
            self.start_button.setText("Измеряется..")
        else:
            self.start_button.setText("Измеряется...")
            self.animation_num = 0
        self.animation_num += 1

    def read_one_func(self):
        # print([child.objectName() for child in self.findChildren(QRadioButton)])
        self.read_one_flag = True
        self.start()
        self.read_one_flag = False
        self.pause()

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

if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = App()
    window.show()
    sys.exit(app.exec())
