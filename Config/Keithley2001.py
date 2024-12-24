"""
Все команды для Keithley2001

:VOLTage[:DC]
    :APERture <n> 
        :AUTO <b>
        :AUTO ONCE
        :AUTO?
    :APERture?
        :NPLCycles <n>
        :AUTO <b>
        :AUTO ONCE
        :AUTO?
    :NPLCycles?
    :RANGe
        [:UPPer] <n>
        [:UPPer]?
        :AUTO <b>
        :AUTO ONCE
            :ULIMit <n>
            :ULIMit?
            :LLIMit <n>
            :LLIMit?
        :AUTO?
    :REFerence <n>
        :STATe <b>
        :STATe?
        :ACQuire
    :REFerence?
    :DIGits <n>
        :AUTO <b>
        :AUTO ONCE
        :AUTO?
    :DIGits?
    :AVERage
        :TCONtrol <name>
        :TCONtrol?
        :COUNt <n>
        :COUNt?
        :ADVanced
            :NTOLerance <n>
            :NTOLerance?
            [:STATe] <b>
            [:STATe]?
        [:STATe] <b>
        [:STATe]?
        :AUTO <b>
        :AUTO ONCE
        :AUTO?
    :FILTer
        [:LPASs]
            [:STATe] <b>
            [:STATe]?

:FRESistance
    :APERture <n>
        :AUTO <b>
        :AUTO ONCE
        :AUTO?
    :APERture?
    :NPLCycles <n>
        :AUTO <b>
        :AUTO ONCE
        :AUTO?
    :NPLCycles?
    :RANGe
        [:UPPer] <n>
        [:UPPer]?
        :AUTO <b>
        :AUTO ONCE
            :ULIMit <n>
            :ULIMit?
            :LLIMit <n>
            :LLIMit?
        :AUTO?
    :REFerence <n>
        :STATe <b>
        :STATe?
        :ACQuire
    :REFerence?
    :DIGits <n>
        :AUTO <b>
        :AUTO ONCE
        :AUTO?
    :DIGits?
    :AVERage
        :TCONtrol <name>
        :TCONtrol?
        :COUNt <n>
        :COUNt?
        :ADVanced
            :NTOLerance <n>
            :NTOLerance?
            [:STATe] <b>
            [:STATe]?
        [:STATe] <b>
        [:STATe]?
        :AUTO <b>
        :AUTO ONCE
        :AUTO?
    :OCOMpensated <b>
    :OCOMpensated?

    !Несколько измерений вариант 1
    :SENS:FUNC 'VOLT:DC'       // Установить функцию измерения DC напряжения
    :SENS:VOLT:NPLC 10         // Задать NPLC
    :TRIG:COUN 10              // Количество измерений = 10
    :TRIG:DEL 0.5              // Задержка перед первым измерением = 0.5 сек
    :TRIG:SOUR TIM             // Использовать таймер для триггера
    :TRIG:TIM 0                // Задержка между последующими измерениями = 0 сек
    :INIT                      // Запустить измерения
    :FETCh?                    // Считать результаты

    !Несколько измерений вариант 2
    :SENS:FUNC 'VOLT:DC'       // Установить функцию измерения DC напряжения
    :SENS:VOLT:NPLC 10         // Задать NPLC

    // Первое измерение с задержкой
    :TRIG:DEL 0.5              // Установить задержку перед первым измерением
    :INIT                      // Запустить измерение
    :FETCh?                    // Считать результат

    // Последующие измерения без задержки
    :TRIG:DEL 0                // Установить задержку = 0
    :INIT                      // Запустить второе измерение
    :FETCh?                    // Считать результат

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

class Keithley2001:
    def __init__(self, app_instance):
        self.app_instance = app_instance
        self.instr = self.app_instance.inst_list
        self.rm = pyvisa.ResourceManager()
        self.instrument = self.rm.open_resource(self.instr["keithley2001"])

        #Задаём параметры #! мб можно убрать
        self.nplc = self.app_instance.NPLC.text()
        self.range12 = self.app_instance.rangeCh12.text() # Если 0, то авто
        self.range34 = self.app_instance.rangeCh34.text()
        self.range56 = self.app_instance.rangeCh56.text()
        self.ch1 = self.app_instance.Ch1.text()
        self.ch2 = self.app_instance.Ch2.text()
        self.ch3 = self.app_instance.Ch3.text()
        self.ch4 = self.app_instance.Ch4.text()
        self.ch5 = self.app_instance.Ch5.text()
        self.ch6 = self.app_instance.Ch6.text()
        self.chterm = self.app_instance.ChTerm.text()

    def reset(self):
        """Сброс настроек прибора"""
        self.instrument.write("*RST")

    def set_dcv_parameters(self, nplc: float,ch: int, range: float, delay: float) -> None:
        """Настройка прибора на измерение постоянного напряжения"""
        self.instrument.write(":SENS:FUNC 'VOLT:DC'")
        self.instrument.write(f":SENS:VOLT:NPLC {nplc}")
        if range == 0:
            self.instrument.write(f":SENS:VOLT:RANG:AUTO ON")
        else:
            self.instrument.write(f":SENS:VOLT:RANG {range}")
        if ch < 10:
            self.instrument.write(f":ROUT:CLOS (@10{ch})")
        else:
            self.instrument.write(f":ROUT:CLOS (@1{ch})")
        self.instrument.write(f":TRIG:DEL {delay}")

    def set_fres_parameters(self, nplc: float,ch: int, range: float, delay: float) -> None:
        """Настройка прибора на измерение 4-проводного сопротивления"""
        self.instrument.write(":SENS:FUNC 'FRES'")
        self.instrument.write(f":SENS:VOLT:NPLC {nplc}")
        if range == 0:
            self.instrument.write(f":SENS:VOLT:RANG:AUTO ON")
        else:
            self.instrument.write(f":SENS:VOLT:RANG {range}")
        if ch < 10:
            self.instrument.write(f":ROUT:CLOS (@10{ch})")
        else:
            self.instrument.write(f":ROUT:CLOS (@1{ch})")
        self.instrument.write(f":TRIG:DEL {delay}")

    def set_res_parameters(self, nplc: float,ch: int, range: float, delay: float) -> None:
        """Настройка прибора на измерение 2-проводного сопротивления"""
        self.instrument.write(":SENS:FUNC 'RES'")
        self.instrument.write(f":SENS:VOLT:NPLC {nplc}")
        if range == 0:
            self.instrument.write(f":SENS:VOLT:RANG:AUTO ON")
        else:
            self.instrument.write(f":SENS:VOLT:RANG {range}")
        if ch < 10:
            self.instrument.write(f":ROUT:CLOS (@10{ch})")
        else:
            self.instrument.write(f":ROUT:CLOS (@1{ch})")
        self.instrument.write(f":TRIG:DEL {delay}")


    def measure(self, meas_count: int) -> list:
        """Запуск измерения и получение результата(должно быть -> list?)"""
        self.instrument.write(f":TRIG:COUN {meas_count}")
        self.instrument.write(":INIT")
        results = []
        for _ in range(meas_count):
            results.append(float(self.instrument.query(":FETCh?")))
        return results


