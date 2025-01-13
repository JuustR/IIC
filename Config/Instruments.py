"""
Добавить опрос по IDN, чтобы было в будующем проще подключать новые приборы
"""

import pyvisa
import time

class InstrumentConnection:
    def __init__(self, app_instance):
        self.rm = pyvisa.ResourceManager()  # Инициализируем ResourceManager
        self.app_instance = app_instance

        self.keithley2010 = None
        self.keithley2000 = None
        self.daq970A = None
        self.rigol = None
        self.keysight = None

        self.E36312A = None
        self.AKIP = None

        self.USB_resources = []
        self.instr_list = {"Smth1": "IDN...1", "Smth2": "IDN...2"}
        self.instr = None # Для проверки всех подключенных приборов

        self.send_IDN = None

    def connect_all(self):
        """Функция для подключения всех выбранных приборов"""
        self.keithley2010_connection()
        # self.keithley2000_connection()
        # self.daq970A_connection()
        # self.keysight_connection()
        # self.rigol_connection()
        # self.E36312A_connection()  # Не работает, т.к. не задан IP
        # self.AKIP_connection()

        # Если нужно будет, чтобы ускорить подключение то можно закоментить
        self.app_instance.ConsolePTE.appendPlainText(
            time.strftime("%H:%M:%S | ", time.localtime()) + 'Подключения к daq, keysight, rigol, E36312A, AKIP закомменчены\n')

        # Должен возвращаться словарь подключенных приборов с их IP/GPIB
        return self.instr_list

    def keithley2010_connection(self):
        try:
            # self.keithley2010 = self.rm.open_resource('GPIB0::' + str(self.app_instance.GPIB.text()) + '::INSTR')
            self.keithley2010 = self.rm.open_resource('GPIB0::' + '16' + '::INSTR')
            self.keithley2010.write("*rst")
            self.app_instance.ConsolePTE.appendPlainText(
                time.strftime("%H:%M:%S | ", time.localtime()) + 'Keithley подключен успешно\n')
            # self.app_instance.statusBar.showMessage('Keithley подключен успешно')
            self.instr_list["keithley2010"] = 'GPIB0::16::INSTR'
        except Exception as e:
            if self.app_instance.show_errors_cb.isChecked():
                self.app_instance.ConsolePTE.appendPlainText(
                    time.strftime("%H:%M:%S | ", time.localtime()) + f'Ошибка подключения Keithley2010!\n'
                                                                     f'GPIB на приборе должен быть 16\n'
                                                                     f'Описание ошибки: {e}\n')
            else:
                print('Ошибка подключения Keithley2010')

    def keithley2000_connection(self):
        try:
            # self.keithley2000 = self.rm.open_resource('GPIB0::' + str(self.app_instance.GPIB.text()) + '::INSTR')
            self.keithley2000 = self.rm.open_resource('GPIB0::' + '16' + '::INSTR')
            self.keithley2000.write("*rst")
            self.app_instance.ConsolePTE.appendPlainText(
                time.strftime("%H:%M:%S | ", time.localtime()) + 'Keithley подключен успешно\n')
            # self.app_instance.statusBar.showMessage('Keithley подключен успешно')
            self.instr_list["keithley2000"] = 'GPIB0::16::INSTR'
        except Exception as e:
            if self.app_instance.show_errors_cb.isChecked():
                self.app_instance.ConsolePTE.appendPlainText(
                    time.strftime("%H:%M:%S | ", time.localtime()) + f'Ошибка подключения Keithley2000!\n'
                                                                     f'GPIB на приборе должен быть 16\n'
                                                                     f'Описание ошибки: {e}\n')
            else:
                print('Ошибка подключения Keithley2000')

    def daq970A_connection(self):
        try:
            # self.daq970A = self.rm.open_resource('TCPIP0::' + str(self.w_root.lineIP_1.text()) + '::inst0::INSTR')
            self.daq970A = self.rm.open_resource('TCPIP0::' + "169.254.50.10" + '::inst0::INSTR')
            self.daq970A.write("*RST")
            self.app_instance.ConsolePTE.appendPlainText(
                time.strftime("%H:%M:%S | ", time.localtime()) + 'DAQ подключен успешно\n')
            # self.app_instance.statusBar.showMessage('Keithley подключен успешно')
            self.instr_list["daq970A"] = 'TCPIP0::169.254.50.10::inst0::INSTR'
        except Exception as e:
            if self.app_instance.show_errors_cb.isChecked():
                self.app_instance.ConsolePTE.appendPlainText(
                    time.strftime("%H:%M:%S | ", time.localtime()) + f'Ошибка подключения DAQ970A!\n'
                                                                     f'IP прибора должен быть 169.254.50.10\n'
                                                                     f'Описание ошибки: {e}\n')
            else:
                print('Ошибка подключения DAQ970A')

    def keysight_connection(self):
        try:
            # self.keysight = self.rm.open_resource('TCPIP0::' + str(self.w_root.lineIP_1.text()) + '::inst0::INSTR')
            self.keysight = self.rm.open_resource('TCPIP0::' + "192.168.0.103" + '::inst0::INSTR')
            self.keysight.write("*RST")
            self.app_instance.ConsolePTE.appendPlainText(
                time.strftime("%H:%M:%S | ", time.localtime()) + 'keysight подключен успешно\n')
            # self.app_instance.statusBar.showMessage('Keithley подключен успешно')
            self.instr_list["keysight"] = 'TCPIP0::192.168.0.103::inst0::INSTR'
        except Exception as e:
            if self.app_instance.show_errors_cb.isChecked():
                self.app_instance.ConsolePTE.appendPlainText(
                    time.strftime("%H:%M:%S | ", time.localtime()) + f'Ошибка подключения keysight!\n'
                                                                     f'IP прибора должен быть 192.168.0.103\n'
                                                                     f'Описание ошибки: {e}\n')
            else:
                print('Ошибка подключения keysight')

    def rigol_connection(self):
        #! Поменять IP
        try:
            # self.rigol = self.rm.open_resource('TCPIP0::' + str(self.w_root.lineIP_1.text()) + '::inst0::INSTR')
            self.rigol = self.rm.open_resource('TCPIP0::' + "169.254.50.9" + '::inst0::INSTR')
            self.rigol.write("*RST")
            self.app_instance.ConsolePTE.appendPlainText(
                time.strftime("%H:%M:%S | ", time.localtime()) + 'Rigol подключен успешно\n')
            self.instr_list["Rigol"] = 'TCPIP0::169.254.50.9::inst0::INSTR'
        except Exception as e:
            if self.app_instance.show_errors_cb.isChecked():
                self.app_instance.ConsolePTE.appendPlainText(
                    time.strftime("%H:%M:%S | ", time.localtime()) + f'Ошибка подключения Rigol!\n'
                                                                     f'IP прибора должен быть 169.254.50.9\n'
                                                                     f'Описание ошибки: {e}\n')
            else:
                print('Ошибка подключения Rigol')

    def E36312A_connection(self):
        #! Нужен IP_BP или что-то такое
        try:
            self.E36312A = self.rm.open_resource('TCPIP0::' + str(self.app_instance.IP_BP.text()) + '::inst0::INSTR')
            self.app_instance.ConsolePTE.appendPlainText(
                    time.strftime("%H:%M:%S | ", time.localtime()) + f'E36312A подключен успешно\n')
        except Exception as e:
            if self.app_instance.show_errors_cb.isChecked():
                self.app_instance.ConsolePTE.appendPlainText(
                    time.strftime("%H:%M:%S | ", time.localtime()) + f'Ошибка подключения E36312A!\n'
                                                                     f'Проверьте IP прибора\n'
                                                                     f'Описание ошибки: {e}\n')
            else:
                print('Ошибка подключения E36312A')

    def AKIP_connection(self):
        # Проверка USB ресурсов
        self.USB_resources = [res for res in self.rm.list_resources() if res.startswith('USB')]
        n = 0
        self.instr_check()
        try:
            while n < len(self.USB_resources):
                try:
                    self.AKIP = self.rm.open_resource(self.USB_resources[n])
                    self.send_IDN = self.AKIP.query("*IDN?")
                except Exception as e:
                    self.send_IDN = None

                if self.send_IDN == 'ITECH Ltd., IT6333A, 800572020767710004, 1.11-1.08\n':
                    self.AKIP.write("*rst")
                    self.app_instance.ConsolePTE.appendPlainText(
                    time.strftime("%H:%M:%S | ", time.localtime()) + f'АКИП подключен успешно\n')
                    break

                n += 1
        except Exception as e:
            if self.app_instance.show_errors_cb.isChecked():
                self.app_instance.ConsolePTE.appendPlainText(
                    time.strftime("%H:%M:%S | ", time.localtime()) + f'Ошибка подключения АКИПа!\n'
                                                                     f'Проверьте IDN\'ы в консоле\n'
                                                                     f'Описание ошибки: {e}\n')
            else:
                print('Ошибка подключения АКИПа')


    def instr_check(self):
        """Поочередный вывод IDN всех приборов по юсб(для отладки)"""
        for i in self.rm.list_resources():
            self.instr = self.rm.open_resource(i)
            print("Для (" + i + ") IDN будет: " + self.instr.query("*IDN?"))
