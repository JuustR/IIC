"""
Класс с настройками и запуском измерений на Keithley 2010

"""

import pyvisa
import time

class Keithley2010:
    def __init__(self, app_instance):
        self.app_instance = app_instance
        self.instr = self.app_instance.inst_list
        self.rm = pyvisa.ResourceManager()
        self.instrument = self.rm.open_resource(self.instr["keithley2010"])

    def reset(self):
        """Сброс настроек прибора"""
        self.instrument.write("*RST")

    def set_dcv_parameters(self, nplc: float,ch: int, range: float, delay: float) -> None:
        """Настройка прибора на измерение постоянного напряжения"""
        self.instrument.write(":SENS:FUNC 'VOLT:DC'")
        self.instrument.write(f":SENS:VOLT:DC:NPLC {nplc}")
        if range == 0:
            self.instrument.write(f":SENS:VOLT:DC:RANG:AUTO ON")
        else:
            self.instrument.write(f":SENS:VOLT:DC:RANG {range}")
        if ch < 10:
            self.instrument.write(f":ROUT:CLOS (@{ch})")
        else:
            self.instrument.write(f":ROUT:CLOS (@{ch})")
        self.instrument.write(f":TRIG:DEL {delay}")
        self.instrument.write("DISP:ENAB ON")

    def set_fres_parameters(self, nplc: float,ch: int, range: float, delay: float) -> None:
        """Настройка прибора на измерение 4-проводного сопротивления"""
        self.instrument.write(":SENS:FUNC 'FRES'")
        self.instrument.write(f":SENS:FRES:NPLC {nplc}")
        if range == 0:
            self.instrument.write(f":SENS:FRES:RANG:AUTO ON")
        else:
            self.instrument.write(f":SENS:FRES:RANG {range}")
        if ch < 10:
            self.instrument.write(f":ROUT:CLOS (@{ch})")
        else:
            self.instrument.write(f":ROUT:CLOS (@{ch})")
        self.instrument.write(f":TRIG:DEL {delay}")
        self.instrument.write("DISP:ENAB ON")

    def set_res_parameters(self, nplc: float,ch: int, range: float, delay: float) -> None:
        """Настройка прибора на измерение 2-проводного сопротивления"""
        self.instrument.write(":SENS:FUNC 'RES'")
        self.instrument.write(f":SENS:RES:NPLC {nplc}")
        if range == 0:
            self.instrument.write(f":SENS:RES:RANG:AUTO ON")
        else:
            self.instrument.write(f":SENS:RES:RANG {range}")
        if ch < 10:
            self.instrument.write(f":ROUT:CLOS (@{ch})")
        else:
            self.instrument.write(f":ROUT:CLOS (@{ch})")
        self.instrument.write(f":TRIG:DEL {delay}")
        self.instrument.write("DISP:ENAB ON")


    def measure(self, meas_count: int) -> list:
        """Запуск измерения и получение результата(должно быть -> list?)"""
        print(time.strftime("%H:%M:%S | ", time.localtime()))
        # self.instrument.write(f":TRIG:COUN {meas_count}")
        # self.instrument.write(":INIT")
        self.instrument.write("DISP:ENAB ON")
        self.instrument.write(":INIT:CONT 1")
        self.instrument.write(":INIT:CONT 0")
        # self.instrument.write(":TRIG:SOUR IMM")
        results = []
        for _ in range(meas_count):
            # results.append(float(self.instrument.query(":FETCh?")))
            values = self.instrument.query_ascii_values(":read?")
            results.append(float(values[0]))
        # self.instrument.write("*RST")
        self.instrument.write("DISP:ENAB ON")
        return results


