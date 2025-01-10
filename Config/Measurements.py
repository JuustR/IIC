"""
Класс содержащий основной цикл измерений

Релюшки:
requests.get('http://192.168.0.101:10500/turn_off_current_switch')
requests.get('http://192.168.0.101:10500/turn_off_sample_switch')
requests.get('http://192.168.0.101:10500/turn_off_heater_current_switch')
"""


import time
import requests

from Config.Keithley2010 import Keithley2010
#from Config.Rigol import Rigol

class Measurements:
    def __init__(self, app_instance):
        self.app_instance = app_instance
        self.settings = self.app_instance.settings_dict
        self.inst_list = self.app_instance.inst_list

        self.fres_value = None

    def cycle(self):
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
        # ТермоЭДС
        requests.get('http://192.168.0.101:10500/turn_on_heater_current_switch')
        #! Тут место для нагревателя)
        


        pass


    def termoemf_step(self):
        """
        N строки
        Время 1
        R термометра 1
        Термопара 1 * N read'ов
        Термопара 2 * N read'ов
        Пустая ячейка
        Мужду термопарами 1 * N read'ов
        Мужду термопарами 2 * N read'ов
        Пустая ячейка
        Ячейки под сопротивление(2*R*(N read'ов) + Катушка + Пустая)
        R термометра 2
        Время 2
        Системное время
        """
        # Номер строки
        number = self.app_instance.start_line_le.text()

        # Время 1
        time1 = time.time() - self.app_instance.start_time

        # R термометра 1
        termometer1 = self.temperature()

        # Термопары
        all_tc = self.termoemf()
        r1 = [i for i in all_tc["ch1"]]
        r2 = [i for i in all_tc["ch2"]]

        # Между
        all_tc = self.termoemf()
        r3 = [i for i in all_tc["ch3"]]
        r4 = [i for i in all_tc["ch4"]]

        # R термометра 2
        termometer2 = self.temperature()

        # Время 2
        time2 = time.time() - self.app_instance.start_time

        # Системное время
        system_time = str(time.strftime("%d.%m.%Y  %H:%M:%S", time.localtime()))


    def resistance_step(self):
        """
        N строки
        Время 1
        R термометра 1
        Пустая Термопара 1 * N read'ов
        Пустая Термопара 2 * N read'ов
        Пустая ячейка
        Пустая Мужду термопарами 1 * N read'ов
        Пустая Мужду термопарами 2 * N read'ов
        Пустая ячейка
        Сопротивление 1 * N read'ов
        Сопротивление 2 * N read'ов
        Катушка
        Пустая ячейка
        R термометра 2
        Время 2
        Системное время
        """
        # Номер строки
        number = self.app_instance.start_line_le.text()

        # Время 1
        time1 = time.time() - self.app_instance.start_time

        # R термометра 1
        termometer1 = self.temperature()

        # Сопротивления
        all_res = self.resistance()
        r5 = [i for i in all_res["ch5"]]
        r6 = [i for i in all_res["ch6"]]

        # Катушка
        kat = self.app_instance.r_cell.text()

        # R термометра 2
        termometer2 = self.temperature()

        # Время 2
        time2 = time.time() - self.app_instance.start_time

        # Системное время
        system_time = str(time.strftime("%d.%m.%Y  %H:%M:%S", time.localtime()))

    def temperature(self):
        """Функция измерения температуры"""
        if self.app_instance.combobox_scan.currentText() == "keithley2010":
            instrument = Keithley2010(self)
            instrument.set_fres_parameters(float(self.settings['nplc_term']),
                                      int(self.settings['ch_term1']),
                                      range=0,
                                      delay=0)
            self.fres_value = instrument.measure(1)
            instrument.reset()  # Сброс настроек перед напряжением
        else:
            self.fres_value = "Error"


        return self.fres_value

    def resistance(self):
        """Функция измерения сопротивления"""
        res_results = {}
        if self.app_instance.combobox_scan.currentText() == "keithley2010":
            instrument = Keithley2010(self)
            for i in range(5, 7):
                ch_line_edit = self.settings[f"ch{i}"]
                delay_line_edit = self.settings[f"delay_ch{i}"]
                range_line_edit = self.settings[f"range_ch{i}"]
                nplc_line_edit = self.settings[f"nplc_ch{i}"]

                instrument.set_dcv_parameters(float(nplc_line_edit),
                                         int(ch_line_edit),
                                         float(range_line_edit),
                                         float(delay_line_edit))  # Первая задержка
                res_results[f"ch{i}"] = instrument.measure(meas_count=1)  # Первое измерение

                # Измерения для больше чем одного read
                if int(self.settings["n_read_ch56"]) > 1:
                    instrument.set_dcv_parameters(float(nplc_line_edit),
                                             int(ch_line_edit),
                                             float(range_line_edit),
                                             delay=0)  # Остальные измерения
                    res_results[f"ch{i}"].extend(
                        instrument.measure(meas_count=(int(self.settings["n_read_ch56"]) - 1)))  # 4 оставшихся измерения
                else:
                    continue
            #! Убрать
            for channel, results in res_results.items():
                print(f"DCV on channel {channel}: {results}")
            instrument.reset()  # Сброс настроек перед сопротивлением
        else:
            res_results["ch5"] = "Error"

        return res_results

    def termoemf(self):
        """Функция измерения термоЭДС"""
        termoemf_results = {}
        if self.app_instance.combobox_scan.currentText() == "keithley2010":
            instrument = Keithley2010(self)
            for i in range(1, 5):
                ch_line_edit = self.settings[f"ch{i}"]
                delay_line_edit = self.settings[f"delay_ch{i}"]
                range_line_edit = self.settings[f"range_ch{i}"]
                nplc_line_edit = self.settings[f"nplc_ch{i}"]

                instrument.set_dcv_parameters(float(nplc_line_edit),
                                         int(ch_line_edit),
                                         float(range_line_edit),
                                         float(delay_line_edit))  # Первая задержка
                termoemf_results[f"ch{i}"] = instrument.measure(meas_count=1)  # Первое измерение

                # Измерения для больше чем одного read
                if i < 3:
                    if int(self.settings["n_read_ch12"]) > 1:
                        instrument.set_dcv_parameters(float(nplc_line_edit),
                                                 int(ch_line_edit),
                                                 float(range_line_edit),
                                                 delay=0)  # Остальные измерения
                        termoemf_results[f"ch{i}"].extend(
                            instrument.measure(meas_count=(int(self.settings["n_read_ch12"]) - 1)))  # 4 оставшихся измерения
                    else:
                        continue
                else:
                    if int(self.settings["n_read_ch34"]) > 1:
                        instrument.set_dcv_parameters(float(nplc_line_edit),
                                                 int(ch_line_edit),
                                                 float(range_line_edit),
                                                 delay=0)  # Остальные измерения
                        termoemf_results[f"ch{i}"].extend(
                            instrument.measure(meas_count=(int(self.settings["n_read_ch34"]) - 1)))  # 4 оставшихся измерения
                    else:
                        continue
            #! Убрать
            for channel, results in termoemf_results.items():
                print(f"DCV on channel {channel}: {results}")
            instrument.reset()  # Сброс настроек перед сопротивлением
        else:
            termoemf_results["ch1"] = "Error"

        return termoemf_results

    def add_dcv(self):
        pass


    # def rigol_measurements(self):
    #     instr = Rigol(self)
    #
    #     # Настройка и измерение 4-проводного сопротивления на канале 101
    #     instr.set_fres_parameters(
    #         float(self.nplc_term.text()),
    #         int(self.ch_term1.text()),
    #         range=0,
    #         delay=0
    #     )
    #     fres_res_1 = instr.measure(1)
    #     print(f"FRES on channel 101: {fres_res_1}")
    #
    #     # Словарь для хранения результатов измерений
    #     dcv_results = {}
    #
    #     # Перебор каналов с параметрами
    #     for i in range(1, 7):
    #         ch_line_edit = self.findChild(QLineEdit, f"ch{i}")  # Поиск элемента с именем ch{i}
    #         delay_line_edit = self.findChild(QLineEdit, f"dealy_ch{i}")
    #         range_line_edit = self.findChild(QLineEdit, f"range_ch{i}")
    #         nplc_line_edit = self.findChild(QLineEdit, f"nplc_ch{i}")
    #
    #
    #         try:
    #             # Открытие канала
    #             instr.open_channel(ch_line_edit)
    #
    #             # Настройка и измерение на текущем канале
    #             instr.set_dcv_parameters(
    #                 float(nplc_line_edit.text()),
    #                 int(ch_line_edit.text()),
    #                 float(range_line_edit.text()),
    #                 float(delay_line_edit.text())
    #             )
    #             dcv_results[f"ch{i}"] = instr.measure(meas_count=1)
    #
    #             # Дополнительные измерения (если требуется)
    #             if i < 3:
    #                 additional_reads = int(self.n_read_ch12.text()) - 1
    #             elif 2 < i < 5:
    #                 additional_reads = int(self.n_read_ch34.text()) - 1
    #             else:
    #                 additional_reads = int(self.n_read_ch56.text()) - 1
    #
    #             if additional_reads > 0:
    #                 instr.set_dcv_parameters(
    #                     float(nplc_line_edit.text()),
    #                     int(ch_line_edit.text()),
    #                     float(range_line_edit.text()),
    #                     delay=0  # Установка задержки для оставшихся измерений
    #                 )
    #                 dcv_results[f"ch{i}"].extend(instr.measure(meas_count=additional_reads))
    #         finally:
    #             # Закрытие канала
    #             instr.close_channel(ch_line_edit)
    #
    #     # Вывод результатов измерений
    #     for channel, results in dcv_results.items():
    #         print(f"DCV on channel {channel}: {results}")
    #
    #     # Повторное измерение 4-проводного сопротивления на канале 101
    #     instr.set_fres_parameters(
    #         float(self.nplc_term.text()),
    #         int(self.ch_term1.text()),
    #         range=0,
    #         delay=0
    #     )
    #     fres_result_2 = instr.measure(1)
    #     print(f"FRES on channel 101 (repeat): {fres_result_2}")
