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
        # self.start_button.clicked.connect(self.on_start_clicked)
        # checkBox'ы
        self.rele_cb.stateChanged.connect(self.rele_cb_clicked)
        self.instr1_cb.stateChanged.connect(self.instr1_cb_clicked)
        self.instr2_cb.stateChanged.connect(self.instr2_cb_clicked)

        # Подключаем кнопки меню настроек
        # self.save_settings_pb.clicked.connect(self.save_settings)

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
        # self.pause_button.clicked.connect(self.pause)
        # self.connect_button.clicked.connect(self.connect)
        # self.read_one.clicked.connect(self.read_one_func)
        # self.save_button.clicked.connect(self.save_settings)

        self.mainthread = None

        # self.load_tab1_settings()

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
        if self.rts_flag:
            self.start_button.setText("Измеряется")
            self.start_flag = True
            self.disable_enable_ch()
            self.animation_timer.start()
            try:
                if self.wb is None:
                    print("Перед запуском убедитесь, что Excel создан")
                    self.statusbar.showMessage("Перед запуском убедитесь, что Excel создан")
                    self.pause()
                    return
                if self.ip_daq_flag:
                    if self.mainthread is None or not self.mainthread.isRunning():
                        self.mainthread = ExcelWriterThread(self)

                        self.mainthread.update_values_signal.connect(self.update_values)
                        self.mainthread.update_excel_signal.connect(self.update_excel)
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
        self.excel_row.setText(f'{num}')

    def pause(self):
        self.start_flag = False
        self.disable_enable_ch()

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

    def save_settings(self):
        for i in range(1, 21):
            checkbox = self.findChild(QCheckBox, f"ch{i}")
            if checkbox:
                self.settings.setValue(f"ch{i}", checkbox.isChecked())

            dcv_radio = self.findChild(QRadioButton, f"dcv{i}")
            if dcv_radio:
                self.settings.setValue(f"dcv{i}", dcv_radio.isChecked())

            fres_radio = self.findChild(QRadioButton, f"fres{i}")
            if fres_radio:
                self.settings.setValue(f"fres{i}", fres_radio.isChecked())

        self.settings.setValue("ip_daq", self.ip_daq.text())
        self.settings.setValue("time_oprosa", self.time_oprosa.text())
        self.settings.setValue("nplc_dcv", self.nplc_dcv.text())
        self.settings.setValue("nplc_fres", self.nplc_fres.text())

        self.statusbar.showMessage('Настройки сохранены')

    def load_settings(self):
        for i in range(1, 21):
            checkbox = self.findChild(QCheckBox, f"ch{i}")
            if checkbox:
                checkbox.setChecked(self.settings.value(f"ch{i}", False, type=bool))

            dcv_radio = self.findChild(QRadioButton, f"dcv{i}")
            fres_radio = self.findChild(QRadioButton, f"fres{i}")

            if dcv_radio:
                dcv_checked = self.settings.value(f"dcv{i}", False, type=bool)
                dcv_radio.blockSignals(True)
                dcv_radio.setChecked(dcv_checked)
                dcv_radio.blockSignals(False)

            if fres_radio:
                fres_checked = self.settings.value(f"fres{i}", False, type=bool)
                fres_radio.blockSignals(True)
                fres_radio.setChecked(fres_checked)
                fres_radio.blockSignals(False)

        self.ip_daq.setText(self.settings.value("ip_daq", ""))
        self.time_oprosa.setText(self.settings.value("time_oprosa", ""))
        self.nplc_dcv.setText(self.settings.value("nplc_dcv", ""))
        self.nplc_fres.setText(self.settings.value("nplc_fres", ""))

    def read_one_func(self):
        # print([child.objectName() for child in self.findChildren(QRadioButton)])
        self.read_one_flag = True
        self.start()
        self.read_one_flag = False
        self.pause()


if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = App()
    window.show()
    sys.exit(app.exec())
