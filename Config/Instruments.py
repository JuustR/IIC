"""
Добавить опрос по IDN, чтобы было в будующем проще подключать новые приборы
"""

import pyvisa
import time

from PyQt6.QtCore import QThread, pyqtSignal


class ConnectionThread(QThread):
    log_signal = pyqtSignal(str)
    result_signal = pyqtSignal(list, list)

    def __init__(self, inst_instance):
        super().__init__()
        self.inst_instance = inst_instance

    def run(self):
        # try:
        #     instr_list, powersource_list = self.inst_instance.connect_all()
        #     self.result_signal.emit(instr_list, powersource_list)
        # except Exception as e:
        #     self.log_signal.emit(f"Ошибка: {e}")
        instr_list, powersource_list = self.inst_instance.connect_all()
        self.result_signal.emit(instr_list, powersource_list)

class InstrumentConnection:
    def __init__(self, app_instance):
        self.rm = pyvisa.ResourceManager()  # Инициализируем ResourceManager
        self.app_instance = app_instance
        self.log_signal = app_instance.log_signal
        # self.formatted_time = self.app_instance.formatted_time

        self.keithley2010 = None
        self.keithley2000 = None
        self.daq970A = None
        self.rigol = None
        self.keysight = None

        self.E36312A = None
        self.AKIP = None

        self.USB_resources = []
        self.instr_list = {"Smth1": "IDN...1", "Smth2": "IDN...2"}
        self.powersource_dict = {"PS1": "IDN"}
        self.instr = None  # Для проверки всех подключенных приборов

        self.send_IDN = None

    def log_message(self, message, exception=None):
        error_message = f"{message}"
        if exception:
            error_message += f"\n{exception}"
        self.log_signal.emit(error_message)
        # self.app_instance.ConsolePTE.appendPlainText(error_message)

    def connect_all(self):
        """Функция для подключения всех выбранных приборов"""
        # self.keithley2010_connection()
        # self.keithley2000_connection()
        # self.daq970A_connection()
        self.keysight_connection()
        self.rigol_connection()
        # self.E36312A_connection()  # Не работает, т.к. не задан IP
        self.akip_connection()

        # Если нужно будет, чтобы ускорить подключение то можно закоментить
        # Все приборы: keithley2010, keithley2000, daq970A, keysight, rigol, E36312A, AKIP
        self.log_message('Подключения к keithley2010, keithley2000, daq970A, E36312A закомменчены')

        # Должен возвращаться словарь подключенных приборов с их IP/GPIB
        return self.instr_list, self.powersource_dict

    def keithley2010_connection(self):
        try:
            # self.keithley2010 = self.rm.open_resource('GPIB0::' + str(self.app_instance.GPIB.text()) + '::INSTR')
            self.keithley2010 = self.rm.open_resource('GPIB0::' + '16' + '::INSTR')
            self.keithley2010.write("*rst")
            self.log_message('Keithley подключен успешно')
            # self.app_instance.statusBar.showMessage('Keithley подключен успешно')
            self.instr_list["keithley2010"] = 'GPIB0::16::INSTR'
        except Exception as e:
            if self.app_instance.show_errors_cb.isChecked():
                self.log_message('Ошибка подключения Keithley2010!\n'
                                 'GPIB на приборе должен быть 16\nОписание ошибки:', e)
            else:
                self.log_message('Ошибка подключения Keithley2010!')

    def keithley2000_connection(self):
        try:
            # self.keithley2000 = self.rm.open_resource('GPIB0::' + str(self.app_instance.GPIB.text()) + '::INSTR')
            self.keithley2000 = self.rm.open_resource('GPIB0::' + '16' + '::INSTR')
            self.keithley2000.write("*rst")
            self.log_message('Keithley подключен успешно')
            self.instr_list["keithley2000"] = 'GPIB0::16::INSTR'
        except Exception as e:
            if self.app_instance.show_errors_cb.isChecked():
                self.log_message('Ошибка подключения Keithley2000!\n'
                                 'GPIB на приборе должен быть 16\nОписание ошибки:', e)
            else:
                self.log_message('Ошибка подключения Keithley2000')

    def daq970A_connection(self):
        try:
            # self.daq970A = self.rm.open_resource('TCPIP0::' + str(self.w_root.lineIP_1.text()) + '::inst0::INSTR')
            self.daq970A = self.rm.open_resource('TCPIP0::' + "169.254.50.10" + '::inst0::INSTR')
            self.daq970A.write("*RST")
            self.log_message('DAQ подключен успешно')
            self.instr_list["daq970A"] = 'TCPIP0::169.254.50.10::inst0::INSTR'
        except Exception as e:
            if self.app_instance.show_errors_cb.isChecked():
                self.log_message('Ошибка подключения DAQ970A!\n'
                                 'IP прибора должен быть 169.254.50.10\nОписание ошибки:', e)
            else:
                self.log_message(f'Ошибка подключения DAQ970A!')

    def keysight_connection(self):
        try:
            # self.keysight = self.rm.open_resource('TCPIP0::' + str(self.w_root.lineIP_1.text()) + '::inst0::INSTR')
            ip = f'TCPIP0::{self.app_instance.ip_rigol.text()}::inst0::INSTR'
            self.keysight = self.rm.open_resource(ip)
            self.keysight.write("*RST")
            self.log_message('keysight подключен успешно')
            self.instr_list["keysight"] = ip
        except Exception as e:
            if self.app_instance.show_errors_cb.isChecked():
                self.log_message('Ошибка подключения keysight!\n'
                                 'IP прибора должен быть 192.168.0.100\nОписание ошибки:', e)
            else:
                self.log_message('Ошибка подключения keysight!')

    def rigol_connection(self):
        # ! Поменять IP
        try:
            # self.rigol = self.rm.open_resource('TCPIP0::' + str(self.w_root.lineIP_1.text()) + '::inst0::INSTR')
            self.rigol = self.rm.open_resource('TCPIP0::' + "169.254.50.9" + '::inst0::INSTR')
            self.rigol.write("*RST")
            self.log_message('Rigol подключен успешно')
            self.instr_list["Rigol"] = 'TCPIP0::169.254.50.9::inst0::INSTR'
        except Exception as e:
            if self.app_instance.show_errors_cb.isChecked():
                self.log_message('Ошибка подключения Rigol!\n'
                                 'IP прибора должен быть 169.254.50.9\nОписание ошибки:', e)
            else:
                self.log_message('Ошибка подключения Rigol!')

    def E36312A_connection(self):
        # ! Нужен IP_BP или что-то такое
        try:
            self.E36312A = self.rm.open_resource('TCPIP0::' + str(self.app_instance.IP_BP.text()) + '::inst0::INSTR')
            self.log_message('E36312A подключен успешно')
            self.powersource_dict["E36312A"] = 'TCPIP0::169.254.50.9::inst0::INSTR'  # !
        except Exception as e:
            if self.app_instance.show_errors_cb.isChecked():
                self.log_message('Ошибка подключения E36312A!\n'
                                 'Проверьте IP прибора\nОписание ошибки:', e)
            else:
                self.log_message('Ошибка подключения E36312A!')

    def akip_connection(self):
        # Проверка USB ресурсов
        self.USB_resources = [res for res in self.rm.list_resources() if res.startswith('USB')]
        # print(self.USB_resources)
        # print(self.rm.list_resources())
        n = 0
        # self.instr_check()  # ожно закоментить
        try:
            while n < len(self.USB_resources):
                try:
                    self.AKIP = self.rm.open_resource(self.USB_resources[n])
                    self.send_IDN = self.AKIP.query("*IDN?")
                    print(self.send_IDN)
                except Exception as e:
                    self.send_IDN = None

                if self.send_IDN == 'ITECH Ltd., IT6333A, 800572020777870001, 1.11-1.08\n':
                    self.AKIP.write("*rst")
                    self.powersource_dict["AKIP"] = self.USB_resources[n]
                    self.log_message('АКИП подключен успешно')
                    break

                n += 1
        except Exception as e:
            if self.app_instance.show_errors_cb.isChecked():
                self.log_message('Ошибка подключения АКИПа!\n'
                                 'Проверьте IDN\'ы в консоле\nОписание ошибки:', e)
            else:
                self.log_message('Ошибка подключения АКИПа!')

    def instr_check(self):
        """Поочередный вывод IDN всех приборов по юсб(для отладки)"""
        for i in self.rm.list_resources():
            self.instr = self.rm.open_resource(i)
            print("Для (" + i + ") IDN будет: " + self.instr.query("*IDN?"))
