"""
    Команды настройки Rigol

    !Вариант для наших измерений
    // Настройка измерений FRES
    :SENS:FUNC 'FRES'            // Установить функцию FRES
    :SENS:FRES:NPLC 10           // NPLC для FRES
    :SENS:FRES:RANG 1E3          // Диапазон для FRES
    :TRIG:DEL 0.5                // Задержка перед первым измерением
    :TRIG:SOUR IMM               // Автоматический запуск измерения
    :TRIG:COUN 1                 // Одно измерение на триггер
    :INIT                        // Запустить FRES
    :FETCh?                      // Считать результат

    // Настройка измерений DCV
    :SENS:FUNC 'VOLT:DC'         // Установить функцию DCV
    :SENS:VOLT:NPLC 1            // NPLC для DCV
    :SENS:VOLT:RANG 10           // Диапазон для DCV

    // Измерение на канале 101
    :ROUT:CLOS (@101)            // Выбрать канал 101
    :TRIG:DEL 0.2                // Задержка перед первым измерением
    :INIT                        // Запустить DCV
    :FETCh?                      // Считать результат

    // Измерение на канале 102
    :ROUT:CLOS (@102)            // Выбрать канал 102
    :TRIG:DEL 0.2                // Задержка перед первым измерением
    :INIT                        // Запустить DCV
    :FETCh?                      // Считать результат

    // Повтор FRES
    :SENS:FUNC 'FRES'            // Вернуться к FRES
    :TRIG:DEL 0.5                // Задержка перед первым измерением
    :INIT                        // Запустить FRES
    :FETCh?                      // Считать результат

"""

import pyvisa
import time


class Rigol:
    def __init__(self, app_instance):
        self.app_instance = app_instance
        self.instr = self.app_instance.inst_list
        self.rm = pyvisa.ResourceManager()
        self.Rigol = self.rm.open_resource(self.instr["Rigol"])
        self.keysight = self.rm.open_resource(self.instr["keysight"])

    def reset(self):
        """Сброс настроек прибора"""
        self.Rigol.write("*RST")
        self.keysight.write("*RST")

    def set_dcv_parameters(self, nplc: float, ch: int, range: float, delay: float) -> None:
        """
        Настройка Rigol на переключение канала и Keysight на измерение постоянного напряжения
        """

        self.keysight.write('CONF:VOLTage')
        self.keysight.write("VOLTage:DC:IMP:AUTO ON")  # High-Z
        self.Rigol.write('INST:DMM OFF')
        if int(ch) > 9:
            self.Rigol.write('ROUT:SCAN (@2' + str(ch) + ')')
        else:
            self.Rigol.write('ROUT:SCAN (@20' + str(ch) + ')')

        self.Rigol.write('ROUT:CHAN:ADV:SOUR BUS')
        self.Rigol.write('INIT')

        self.keysight.write(f"VOLTage:DC:NPLC {nplc}")
        if range == 0:
            self.keysight.write("VOLTage:DC:RANGe:AUTO ON")
        else:
            self.keysight.write(f"VOLTage:DC:RANGe {range}")
        self.keysight.write('SYST:LOC')
        time.sleep(delay)

    def set_fres_parameters(self, nplc: float, ch: int, range: float, delay: float) -> None:
        """
        Настройка Rigol на переключение канала и Keysight на измерение 4-проводного сопротивления
        """

        self.Rigol.write('INST:DMM OFF')
        if int(ch) > 9:
            self.Rigol.write('ROUT:SCAN (@2' + str(ch) + ')')
            self.Rigol.write('ROUT:CHAN:FWIR ON,(@2' + str(ch) + ')')
        else:
            self.Rigol.write('ROUT:SCAN (@20' + str(ch) + ')')
            self.Rigol.write('ROUT:CHAN:FWIR ON,(@20' + str(ch) + ')')

        self.Rigol.write('ROUT:CHAN:ADV:SOUR BUS')
        self.Rigol.write('INIT')

        self.keysight.write('CONF:FRES')
        self.keysight.write(f"FRES:NPLC {nplc}")
        if range == 0:
            self.keysight.write("FRES:RANGe:AUTO ON")
        else:
            self.keysight.write(f"FRES:RANGe {range}")
        self.keysight.write('SYST:LOC')
        time.sleep(delay)


    def set_res_parameters(self, nplc: float, ch: int, range: float, delay: float) -> None:
        """
        Настройка Rigol на переключение канала и Keysight на измерение 2-проводного сопротивления
        """
        # ! НЕ РАБОТАЕТ
        nplc = int(nplc) if nplc.is_integer() else float(nplc)
        range = int(range) if range.is_integer() else float(range)
        delay = int(delay) if delay.is_integer() else float(delay)

        self.keysight.write(":SENS:FUNC 'RES'")
        self.keysight.write(f":SENS:RES:NPLC {nplc}")
        if range == 0:
            self.keysight.write(":SENS:RES:RANG:AUTO ON")
        else:
            self.keysight.write(f":SENS:RES:RANG {range}")
        self.keysight.write(f":TRIG:DEL {delay}")

    def measure(self, meas_count: int) -> list:
        """
        Запуск измерений и получение результатов с Keysight
        """
        results = []
        for _ in range(meas_count):
            one_read = self.keysight.query_ascii_values(":READ?")
            results.append(float(one_read[0]))
        # self.Rigol.write('*TRG')
        return results

    def trig_rigol(self):
        self.Rigol.write('*TRG')

    def open_channel(self, ch: int) -> None:
        rigol_channel = f"10{ch}" if ch < 10 else f"1{ch}"
        self.Rigol.write(f":ROUT:CHAN {rigol_channel}, ON")

    def close_channel(self, ch: int) -> None:
        rigol_channel = f"10{ch}" if ch < 10 else f"1{ch}"
        self.Rigol.write(f":ROUT:CHAN {rigol_channel}, OFF")
