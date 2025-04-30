import pyvisa
import time


class DAQ970A:
    def __init__(self, app_instance):
        self.app_instance = app_instance
        self.instr = self.app_instance.inst_list
        self.rm = pyvisa.ResourceManager()
        self.keysight = self.rm.open_resource(self.instr["daq970A"])
        self.keysight.timeout = 5000
        self.nplc_list = [0.001, 0.002, 0.006, 0.02, 0.06, 0.2, 1, 2, 10, 20, 100, 200]
        self.range_dcv_list = [0.1, 1, 10, 100, 300]
        self.range_fres_list = [100, 1000, 10_000, 100_000, 1_000_000, 10_000_000, 100_000_000, 1_000_000_000]

    def reset(self):
        """Сброс настроек прибора"""
        self.keysight.write("*RST")

    def set_dcv_parameters(self, nplc: float, ch: int, range: float, delay: float) -> None:
        """
        Настройка Rigol на переключение канала и Keysight на измерение постоянного напряжения
        """
        # Проверка на правильность записи параметров
        range = range if range in self.range_dcv_list else 0
        nplc = nplc if nplc in self.nplc_list else 1

        if int(ch) > 9:
            if range == 0:
                self.keysight.write(f'CONF:VOLT:DC (@1{ch})')
            else:
                self.keysight.write(f'CONF:VOLT:DC {range}, (@1{ch})')
        else:
            if range == 0:  # Auto range
                self.keysight.write(f'CONF:VOLT:DC (@10{ch})')
            else:
                self.keysight.write(f'CONF:VOLT:DC {range}, (@10{ch})')
        self.keysight.write("VOLT:DC:IMP:AUTO ON")  # High-Z

        self.keysight.write(f"VOLT:DC:NPLC {nplc}")
        # self.keysight.write('SYST:LOC')  # Хз нужно или нет(как будто нет)
        time.sleep(delay)

    def set_fres_parameters(self, nplc: float, ch: int, range: float, delay: float) -> None:
        """
        Настройка Rigol на переключение канала и Keysight на измерение 4-проводного сопротивления
        """
        # Проверка на правильность записи параметров
        range = range if range in self.range_fres_list else 0
        nplc = nplc if nplc in self.nplc_list else 1

        if int(ch) > 9:
            if range == 0:
                self.keysight.write(f'CONF:FRES (@1{ch})')
            else:
                self.keysight.write(f'CONF:FRES {range}, (@1{ch})')
        else:
            if range == 0:  # Auto range
                self.keysight.write(f'CONF:FRES (@10{ch})')
            else:
                self.keysight.write(f'CONF:FRES {range}, (@10{ch})')

        self.keysight.write(f"FRES:NPLC {nplc}")
        # self.keysight.write('SYST:LOC')  # Хз нужно или нет(как будто нет)
        time.sleep(delay)


    def set_res_parameters(self, nplc: float, ch: int, range: float, delay: float) -> None:
        """
        Настройка Rigol на переключение канала и Keysight на измерение 2-проводного сопротивления
        """
        # Проверка на правильность записи параметров
        range = range if range in self.range_fres_list else 0
        nplc = nplc if nplc in self.nplc_list else 1

        if int(ch) > 9:
            if range == 0:
                self.keysight.write(f'CONF:RES (@1{ch})')
            else:
                self.keysight.write(f'CONF:RES {range}, (@1{ch})')
        else:
            if range == 0:  # Auto range
                self.keysight.write(f'CONF:RES (@10{ch})')
            else:
                self.keysight.write(f'CONF:RES {range}, (@10{ch})')

        self.keysight.write(f"RES:NPLC {nplc}")
        # self.keysight.write('SYST:LOC')  # Хз нужно или нет(как будто нет)
        time.sleep(delay)

    def measure(self, meas_count: int) -> list:
        """
        Запуск измерений и получение результатов с Keysight
        """
        results = []
        for _ in range(meas_count):
            one_read = self.keysight.query_ascii_values(":READ?")
            results.append(float(one_read[0]))
        return results
