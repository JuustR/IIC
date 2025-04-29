"""

"""

import time

from PyQt6.QtWidgets import QRadioButton
from PyQt6.uic import loadUi
from PyQt6.QtWidgets import QApplication, QMainWindow, QMessageBox, QPushButton, QCheckBox, QLineEdit
from PyQt6.QtCore import pyqtSignal, QThread, QSettings, QTimer


class ExcelWriterThread(QThread):
    stop_signal = pyqtSignal()
    update_excel_signal = pyqtSignal(int, int, object)
    update_values_signal = pyqtSignal(int)

    def __init__(self, app_instance, parent=None):
        super().__init__(parent)
        self.app_instance = app_instance

        self.daq = self.app_instance.daq
        self.nplc_dcv = self.app_instance.nplc_dcv.text()
        self.nplc_fres = self.app_instance.nplc_fres.text()
        self.read_one_flag = self.app_instance.read_one_flag
        self.statusbar = self.app_instance.statusbar

        self.running = True
        self.excel_row = int(self.app_instance.excel_row.text())  # Начальная строка записи

    def run(self):
        current_row = self.excel_row

        while self.running:
            read_number = 1  # Для перебора каналов
            self.update_excel_signal.emit(current_row, 1, time.time() - self.app_instance.start_time)

            self.statusbar.showMessage("Старт")

            for i in range(1, 21):
                # ! Не знаю, будет ли грузить GUI, если да, то нужно будет что-то придумывать
                checkbox = self.app_instance.findChild(QCheckBox, f"ch{i}")
                if checkbox.isChecked():
                    dcv_radio = self.app_instance.findChild(QRadioButton, f"dcv{i}")
                    # fres_radio = self.app_instance.findChild(QRadioButton, f"fres{i}")
                    if dcv_radio.isChecked():
                        read = self.dcv_read(i, self.nplc_dcv)
                    else:
                        read = self.fres_read(i, self.nplc_fres)
                else:
                    continue

                self.update_excel_signal.emit(current_row, read_number + 1, read)

                read_number += 1

            self.update_excel_signal.emit(current_row, read_number + 1,
                                          str(time.strftime("%d.%m.%Y  %H:%M:%S", time.localtime())))

            current_row += 1  # Переход на следующую строку после записи данных
            self.update_values_signal.emit(current_row)

            if self.read_one_flag:
                self.stop()
            else:
                # Было
                # time_sleep = int(self.app_instance.time_oprosa.text())
                # time.sleep(time_sleep)

                # Стало
                time_sleep = int(self.app_instance.time_oprosa.text())
                while time_sleep > 0:
                    # print(time_sleep)
                    if time_sleep > 10:
                        self.statusbar.showMessage(f"Измерения продолжатся через {time_sleep} секунд")
                    if self.running:
                        if (time_sleep - 1) > 0:
                            time.sleep(1)  # Пауза между записями
                            time_sleep -= 1
                        else:
                            time.sleep(time_sleep)
                            time_sleep -= 1
                            break
                    else:
                        self.stop()
                        break

    def stop(self):
        self.running = False
        self.stop_signal.emit()

    def dcv_read(self, ch, nplc):
        if ch > 9:
            self.daq.write(f"CONF:VOLT:DC (@1{ch})")
        else:
            self.daq.write(f"CONF:VOLT:DC (@10{ch})")
        self.daq.write("VOLT:DC:IMP:AUTO ON")
        self.daq.write(f"VOLT:DC:NPLC {nplc}")
        temp_values = self.daq.query_ascii_values("READ?")
        return temp_values[0]

    def fres_read(self, ch, nplc):
        if ch > 9:
            self.daq.write(f"CONF:FRES (@1{ch})")
        else:
            self.daq.write(f"CONF:FRES (@10{ch})")
        self.daq.write(f"FRES:NPLC {nplc}")
        temp_values = self.daq.query_ascii_values("READ?")
        return temp_values[0]