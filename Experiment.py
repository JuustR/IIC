from GUI import App
import pyvisa as visa


class Experiment(App):
    def __init__(self, app_instance):
        self.app_instance = app_instance

        rm = visa.ResourceManager
        self.keithley = rm.open_resource('GPIB0::16::INSTR')
        self.keithley.write('*rst')
        # self.keithley.write('*rst; status:preset; *cls') # Полная очистка параметров и бит регистра


        # Test keithley
        interval_in_ms = 500
        number_of_readings = 10
        self.keithley.write("status:measurement:enable 512; *sre 1") # устанавливает бит 9 в регистре статуса измерений, что может использоваться для мониторинга определенного события или состояния
        self.keithley.write("sample:count %d" % number_of_readings) # Эта команда устанавливает общее количество измерений
        self.keithley.write("trigger:source bus") # устанавливает источник триггера на системную шину
        self.keithley.write("trigger:delay %f" % (interval_in_ms / 1000.0)) # задает задержку между триггерами в секундах
        self.keithley.write("trace:points %d" % number_of_readings) # устанавливает количество точек данных для записи
        self.keithley.write("trace:feed sense1; feed:control next") # определяет канал, с которого будут записываться данные (sense1), и управляет их передачей на следующий этап обработки

        self.keithley.write("initiate") # make the instrument waiting for a trigger pulse,
        self.keithley.assert_trigger() # trigger it,
        self.keithley.wait_for_srq() # and wait until it sends a “service request”

        voltages = self.keithley.query_ascii_values("trace:data?")
        print("Average voltage: ", sum(voltages) / len(voltages))

        self.keithley.query("status:measurement?") # reset the instrument’s data buffer and SRQ status register
        self.keithley.write("trace:clear; feed:control next")

        # Вывод результатов можно реализовать как:

        # voltages = self.keithley.query_ascii_values("trace:data?") # Список значений
        # print("Average voltage: ", sum(voltages) / len(voltages)) # вывод среднего значения

        # result = self.keithley.query(":fetch?") # Нужно тестить, не до конца понимаю как будет усредняться и будет ли результат в 10чном виде
        # print("Average voltage: ", result)

    def instr_init(self):
        pass

    def fres_meas(self):
        self.keithley.write(":rout:clos (@" + "10" + ")") # Выставляем 10 канал например
        self.keithley.write("sense:function 'resistance'")  # Установить функцию измерения на сопротивление
        self.keithley.write("sense:resistance:mode four")  # Включить 4-х контактный режим измерения
        self.keithley.write("sense:resistance:range:" + "auto on")  # Включить автоматический выбор диапазона
        self.keithley.write("sample:count 1")  # Установить количество выборок на 1 для одного измерения
        self.keithley.write("trigger:source bus")  # Установить источник триггера на системную шину
        self.keithley.write("trigger:delay 0")  # Без задержки триггера, тк одно измерение
        self.keithley.write(":init:cont 0")  # Отключить непрерывную инициализацию

        self.keithley.write("initiate")  # Запуск измерения (Можно убрать и добавить в порядок измерений)
        self.keithley.write(":init:cont 0")  # Отключить непрерывную инициализацию

    def DCvolt_meas(self):
        self.keithley.write(":rout:clos (@" + "10" + ")") # Выставляем 10 канал например
        self.keithley.write("sense:function 'voltage:dc'")  # Установить функцию измерения на постоянное напряжение
        self.keithley.write("sense:voltage:nplc " + "10")  # Установить NPLC на 10
        self.keithley.write("sample:count " + "2")  # Установить количество выборок на 2 для усреднения
        self.keithley.write("trigger:source bus")  # Установить источник триггера на системную шину
        self.keithley.write("trigger:delay " + "0")  # Задержка триггера 0
        self.keithley.write("trace:points " + "2")  # Установить количество точек данных для записи на 2
        self.keithley.write("trace:feed sense1; feed:control next")  # Установить канал для записи данных
        self.keithley.write(":init:cont 1")  # Включить непрерывную инициализацию

        self.keithley.write("initiate")  # Запуск измерения (Можно убрать и добавить в порядок измерений)
        self.keithley.write(":init:cont 0")  # Отключить непрерывную инициализацию
