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
"""

import pyvisa

class Keithley2001:
    def __init__(self, app_instance):
        self.app_instance = app_instance
        self.instr = self.app_instance.inst_list
        self.rm = pyvisa.ResourceManager()
        self.instrument = self.rm.open_resource(self.instr["keithley2001"])

        # Установка параметров при инициализации
        self.set_nplc(app_instance.get("nplc", 1))
        self.set_delay(app_instance.get("delay", 0))
        self.set_readings(app_instance.get("readings", 1))
        self.set_channel(app_instance.get("channel", 1))

    def reset(self):
        """Сброс настроек прибора."""
        self.instrument.write("*RST")

    def set_dcv(self):
        """Настройка прибора на измерение постоянного напряжения (DCV)."""
        self.instrument.write("FUNC 'VOLT:DC'")

    def set_fres(self):
        """Настройка прибора на измерение 4-проводного сопротивления (FRES)."""
        self.instrument.write("FUNC 'FRES'")

    def set_res(self):
        """Настройка прибора на измерение 2-проводного сопротивления (RES)."""
        self.instrument.write("FUNC 'RES'")

    def set_nplc(self, nplc):
        """Установка параметра NPLC (число периодов линий питания).

        Args:
            nplc (float): Значение NPLC.
        """
        self.instrument.write(f"VOLT:NPLC {nplc}")

    def set_delay(self, delay):
        """Установка задержки между измерениями.

        Args:
            delay (float): Значение задержки в секундах.
        """
        self.instrument.write(f"TRIG:DEL {delay}")

    def set_readings(self, count):
        """Установка количества чтений.

        Args:
            count (int): Количество чтений.
        """
        self.instrument.write(f"SAMP:COUN {count}")

    def set_channel(self, channel):
        """Выбор канала измерения (для многоканальных систем).

        Args:
            channel (int): Номер канала.
        """
        self.instrument.write(f"ROUT:CHAN {channel}")

    def measure(self):
        """Запуск измерения и получение результата.

        Returns:
            float: Результат измерения.
        """
        return float(self.instrument.query("READ?"))

    def close(self):
        """Закрытие соединения с прибором."""
        self.instrument.close()
        self.rm.close()


# Пример использования:
if __name__ == "__main__":
    app_instance = {
        "resource_name": "GPIB::24::INSTR",
        "nplc": 10,
        "delay": 0.1,
        "readings": 5,
        "channel": 1
    }
    keithley = Keithley2001(app_instance)
    try:
        keithley.reset()
        keithley.set_dcv()
        result = keithley.measure()
        print(f"Результат измерения DCV: {result} В")
    finally:
        keithley.close()
