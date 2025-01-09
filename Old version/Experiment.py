"""
План по измерениям

1. Проверка подключения к приборам
2. Запись начальных условий
3. Старт измерений
    3.1. Измерение сопротивления
        3.1.1. Замыкание релюшек на ток
        3.1.2. Подача напряжения на токовые контакты
        3.1.3. Запись измеренных данных
        3.1.4. Переключение направления тока (мб секундную задержку между переключениями)
        3.1.5. Повторение пунктов 2-3
        3.1.6. Если измерений >1, то повторяем 4-2-3-4-2-3 умноженное на кол-во измерений (4-2-3 можно в одну ф-ю)
        3.1.7. Размыкание токовых релюшек, отключение нагревателя и пауза (если нужно)
    3.2. Измерения термоЭДС (если нужно)
        3.2.1. Замыкание релюшки на нагреватель
        3.2.2. Подаём напряжение на нагреватель
        3.2.3. Измеряем и записываем n точек полуцикла нагрева
        3.2.4. Размыкаем релюшку нагревателя и отключаем нагреватель
        3.2.5. Измеряем и записываем m точек полуцикла охлаждения
        3.2.6. Выставляем паузу, если нужно
4. Остановка измерений и сохранение файла
"""

import pyvisa as visa


class Experiment:
    def __init__(self, app_instance):
        self.app_instance = app_instance

        rm = visa.ResourceManager
        self.keithley = rm.open_resource('GPIB0::16::INSTR')
        self.keithley.write('*rst')
        # self.keithley2010.write('*rst; status:preset; *cls') # Полная очистка параметров и бит регистра

        # Test keithley2010
        interval_in_ms = 500
        number_of_readings = 10
        self.keithley.write(
            "status:measurement:enable 512; *sre 1")  # устанавливает бит 9 в регистре статуса измерений, что может использоваться для мониторинга определенного события или состояния
        self.keithley.write(
            "sample:count %d" % number_of_readings)  # Эта команда устанавливает общее количество измерений
        self.keithley.write("trigger:source bus")  # устанавливает источник триггера на системную шину
        self.keithley.write(
            "trigger:delay %f" % (interval_in_ms / 1000.0))  # задает задержку между триггерами в секундах
        self.keithley.write("trace:points %d" % number_of_readings)  # устанавливает количество точек данных для записи
        self.keithley.write(
            "trace:feed sense1; feed:control next")  # определяет канал, с которого будут записываться данные (sense1), и управляет их передачей на следующий этап обработки

        self.keithley.write("initiate")  # make the instrument waiting for a trigger pulse,
        self.keithley.assert_trigger()  # trigger it,
        self.keithley.wait_for_srq()  # and wait until it sends a “service request”
 
        voltages = self.keithley.query_ascii_values("trace:data?")
        print("Average voltage: ", sum(voltages) / len(voltages))

        self.keithley.query("status:measurement?")  # reset the instrument’s data buffer and SRQ status register
        self.keithley.write("trace:clear; feed:control next")

        # Вывод результатов можно реализовать как:

        # voltages = self.keithley2010.query_ascii_values("trace:data?") # Список значений
        # print("Average voltage: ", sum(voltages) / len(voltages)) # вывод среднего значения

        # result = self.keithley2010.query(":fetch?") # Нужно тестить, не до конца понимаю как будет усредняться и будет ли результат в 10чном виде
        # print("Average voltage: ", result)

    def instr_init(self):
        pass

    def fres_meas_keithley(self):
        self.keithley.write(":rout:clos (@" + "10" + ")")  # Выставляем 10 канал например
        self.keithley.write("sense:function 'resistance'")  # Установить функцию измерения на сопротивление
        self.keithley.write("sense:resistance:mode four")  # Включить 4-х контактный режим измерения
        self.keithley.write("sense:resistance:range:" + "auto on")  # Включить автоматический выбор диапазона
        self.keithley.write("sample:count 1")  # Установить количество выборок на 1 для одного измерения
        self.keithley.write("trigger:source bus")  # Установить источник триггера на системную шину
        self.keithley.write("trigger:delay 0")  # Без задержки триггера, тк одно измерение
        self.keithley.write(":init:cont 0")  # Отключить непрерывную инициализацию

        self.keithley.write("initiate")  # Запуск измерения (Можно убрать и добавить в порядок измерений)
        self.keithley.write(":init:cont 0")  # Отключить непрерывную инициализацию

    def DC_meas_keithley(self):
        self.keithley.write(":rout:clos (@" + "10" + ")")  # Выставляем 10 канал например
        self.keithley.write("sense:function 'voltage:dc'")  # Установить функцию измерения на постоянное напряжение
        self.keithley.write("sense:voltage:nplc " + "10")  # Установить NPLC на 10
        self.keithley.write("sense:voltage:range:" + "auto on")  # Установить range на auto
        self.keithley.write("sample:count " + "2")  # Установить количество выборок на 2 для усреднения
        self.keithley.write("trigger:source bus")  # Установить источник триггера на системную шину
        self.keithley.write("trigger:delay " + "0")  # Задержка триггера 0
        self.keithley.write("trace:points " + "2")  # Установить количество точек данных для записи на 2
        self.keithley.write("trace:feed sense1; feed:control next")  # Установить канал для записи данных
        self.keithley.write(":init:cont 1")  # Включить непрерывную инициализацию

        self.keithley.write("initiate")  # Запуск измерения (Можно убрать и добавить в порядок измерений)
        self.keithley.write(":init:cont 0")  # Отключить непрерывную инициализацию

    def measur(self):

        pass

    def Instr_connection(self):
        "Функция для подключения приборов"
        rm = visa.ResourceManager()
        
        # PowerSource_allowed - стоит галочка на источник питания
        if self.PowerSource_allowed == True:
            #Разобраться в подключении к USB
            n = 0
            while n < len(self.USB_resources):
                try:
                    AKIP = rm.open_resource(self.USB_resources[n])
                    self.send_IDN = AKIP.query("*IDN?")
                except:
                    self.send_IDN = 'None'
                if self.send_IDN == 'ITECH Ltd., IT6333A, 800572020767710004, 1.11-1.08\n':
                    # print('Проверка условия прошла')
                    AKIP.write("*rst")
                    # print('Команда подана')
                    self.w_root.statusBar.showMessage('Подключение успешно')
                    break
                # print('while1:' + str(n))

                n += 1
            else:
                error = QMessageBox()
                error.setWindowTitle('Не удалось подключиться к АКИП-1142/3')
                error.setText('Убедитесь, что АКИП-1142/3 подключен к компуктеру по USB')
                error.setIcon(QMessageBox.Warning)
                error.setStandardButtons(QMessageBox.Ok)
                error.exec_()
                self.w_root.statusBar.showMessage('Подключение не удалось')

        AKIP.write('INSTrument:NSELect ' + self.heater_channel)
        AKIP.write('APPL CH' + self.heater_channel + ',' + str(self.heater_voltage) + ',1')  # APPL канал, напряжение, ток
        AKIP.write('CHANnel:OUTPut 0')  # + str(int(self.Change_Volt))

        print("Список подключенных приборов: " + rm.list_resources())
        return rm
    
    def Settings(self):
        """Сюда прописываются и хранятся параметры которые измеряются right now"""
        self.akip_channel = 1 # 1-3
        self.akip_voltage = 0 # 0-60 для 1-2 ch и 0-5 для 3 ch
        # self.


        pass