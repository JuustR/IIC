"""
Добавить опрос по IDN, чтобы было в будующем проще подключать новые приборы
"""

import pyvisa

class InstrumentConnection:
    def __init__(self, w_root):
        self.rm = pyvisa.ResourceManager()  # Инициализируем ResourceManager
        self.w_root = w_root
        self.keithley = None
        self.E36312A = None
        self.AKIP = None
        self.USB_resources = []

    def Connect_all(self):
        "Функция для подключения всех выбранных приборов"
        if self.Keithley_allowed:
            "Добавить галочку или выбор сканера к которому подключаться"
            self.Keithley_connection()
        if self.BP_allowed and self.E36312A_allowed:
            "Если есть нагреватель и подключен Е36312А"
            self.E36312A_connection()
        if self.BP_allowed and self.AKIP_allowed:
            "Если есть нагреватель и подключен Е36312А"
            self.AKIP_connection()


    def Keithley_connection(self):
        try:
            self.keithley = self.rm.open_resource('GPIB0::' + str(self.w_root.GPIB.text()) + '::INSTR')
            self.keithley.write("*rst")
            self.w_root.statusBar.showMessage('Keithley подключен успешно')
        except Exception as e:
            self.w_root.statusBar.showMessage(f'Ошибка подключения Keithley: {e}')

    def E36312A_connection(self):

        try:
            self.E36312A = self.rm.open_resource('TCPIP0::' + str(self.w_root.IP_BP.text()) + '::inst0::INSTR')
            self.w_root.statusBar.showMessage('E36312A подключен успешно')
        except Exception as e:
            self.w_root.statusBar.showMessage(f'Ошибка подключения E36312A: {e}')

    def AKIP_connection(self):
        # Проверка USB ресурсов
        self.USB_resources = [res for res in self.rm.list_resources() if res.startswith('USB')]
        n = 0
        try:
            while n < len(self.USB_resources):
                try:
                    self.AKIP = self.rm.open_resource(self.USB_resources[n])
                    self.send_IDN = self.AKIP.query("*IDN?")
                except:
                    self.send_IDN = None

                if self.send_IDN == 'ITECH Ltd., IT6333A, 800572020767710004, 1.11-1.08\n':
                    self.AKIP.write("*rst")
                    self.w_root.statusBar.showMessage('AKIP подключен успешно')
                    break

                n += 1
        except Exception as e:
            self.w_root.statusBar.showMessage(f'AKIP не подключен(или не обнаружается) по USB: {e}')


    def Instr_check(self):
        for i in self.rm.list_resources():
            self.instr = self.rm.open_resource(i)
            print("Для (" + i + ") IDN будет: " + self.instr.query("*IDN?"))