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

class Rigol:
    def __init__(self, app_instance):
        self.app_instance = app_instance
        self.instr = self.app_instance.inst_dict
        self.rm = pyvisa.ResourceManager()
        self.instrument = self.rm.open_resource(self.instr["rigol"])
        self.additional_inst = self.rm.open_resource(self.instr["keysight"])


    def reset(self):
        """Сброс настроек прибора"""
        self.instrument.write("*RST")

    def set_dcv_parameters(self, nplc: float, ch: int, range: float, delay: float) -> None:
        """Настройка Rigol на переключение канала и Keysight на измерение постоянного напряжения"""
        # Настройка Keysight для измерения постоянного напряжения
        self.additional_inst.write(":SENS:FUNC 'VOLT:DC'")
        self.additional_inst.write(f":SENS:VOLT:NPLC {nplc}")
        if range == 0:
            self.additional_inst.write(":SENS:VOLT:RANG:AUTO ON")
        else:
            self.additional_inst.write(f":SENS:VOLT:RANG {range}")
        self.additional_inst.write(f":TRIG:DEL {delay}")

    def set_fres_parameters(self, nplc: float, ch: int, range: float, delay: float) -> None:
        """Настройка Rigol на переключение канала и Keysight на измерение 4-проводного сопротивления"""
        # Настройка Keysight для измерения 4-проводного сопротивления
        self.additional_inst.write(":SENS:FUNC 'FRES'")
        self.additional_inst.write(f":SENS:VOLT:NPLC {nplc}")
        if range == 0:
            self.additional_inst.write(":SENS:RES:RANG:AUTO ON")
        else:
            self.additional_inst.write(f":SENS:RES:RANG {range}")
        self.additional_inst.write(f":TRIG:DEL {delay}")


    def set_res_parameters(self, nplc: float, ch: int, range: float, delay: float) -> None:
        """Настройка Rigol на переключение канала и Keysight на измерение 2-проводного сопротивления"""
        # Настройка Keysight для измерения 2-проводного сопротивления
        self.additional_inst.write(":SENS:FUNC 'RES'")
        self.additional_inst.write(f":SENS:VOLT:NPLC {nplc}")
        if range == 0:
            self.additional_inst.write(":SENS:RES:RANG:AUTO ON")
        else:
            self.additional_inst.write(f":SENS:RES:RANG {range}")
        self.additional_inst.write(f":TRIG:DEL {delay}")

    def measure(self, meas_count: int) -> list:
        """Запуск измерений и получение результатов с Keysight"""
        # Настройка Keysight на количество измерений
        self.additional_inst.write(f":TRIG:COUN {meas_count}")
        self.additional_inst.write(":INIT")

        # Сбор результатов измерений
        results = []
        for _ in range(meas_count):
            results.append(float(self.additional_inst.query(":FETCh?")))

        return results

    def open_channel(self, ch: int) -> None:
        rigol_channel = f"10{ch}" if ch < 10 else f"1{ch}"
        self.instrument.write(f":ROUT:CHAN {rigol_channel}, ON")

    def close_channel(self, ch: int) -> None:
        rigol_channel = f"10{ch}" if ch < 10 else f"1{ch}"
        self.instrument.write(f":ROUT:CHAN {rigol_channel}, OFF")


