"""
Класс содержащий основной цикл измерений

"""

import time
import requests
import pyvisa

from Config.Keithley2010 import Keithley2010
from Config.Rigol import Rigol


class Measurements:
    def __init__(self, app_instance):
        self.rm = pyvisa.ResourceManager()  # Инициализируем ResourceManager

        self.app_instance = app_instance
        self.settings = self.app_instance.settings_dict
        self.inst_list = self.app_instance.inst_list
        self.powersource_list = self.app_instance.powersource_list
        self.formatted_time = self.app_instance.formatted_time

        # Инициализация ИП
        self.AKIP = self.rm.open_resource(self.powersource_list["AKIP"])

        self.change_volt_flag = False  # Флаг отвечающий за переключение направления тока

        self.fres_value = None
        self.number = 11
        self.meas_number = 1

    def toggle_relay(self, relay_type, state):
        """Управление релюшками, где current - направление тока, sample - ток, heater - нагреватель"""
        d = {
            "current": "current_switch",
            "sample": "sample_switch",
            "heater": "heater_current_switch",
        }
        try:
            action = d[relay_type]
            requests.get(f'http://192.168.0.101:10500/turn_{state}_{action}')
        except Exception as e:
            self.log_message(f"Ошибка переключения реле {relay_type} ({state})", e)

    def control_heater(self, channel, voltage, state):
        """Управление нагревателем"""
        try:
            if self.app_instance.combobox_power.currentText() == "AKIP":
                self.AKIP.write(f'INSTrument:NSELect {channel}')
                self.AKIP.write(f'APPL CH{channel},{voltage},1')
                self.AKIP.write(f'CHANnel:OUTPut {1 if state == "on" else 0}')
        except Exception as e:
            self.log_message(f"Ошибка управления нагревателем ({state})", e)

    def log_message(self, message, exception=None):
        error_message = f"{self.formatted_time}{message}\n"
        if exception:
            error_message += f"{exception}\n"
        self.app_instance.ConsolePTE.appendPlainText(error_message)

    def cycle_S_R(self):
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
        # Проверка обновления настроек
        try:
            if self.app_instance.settings_changed_flag:
                self.app_instance.copy_settings_to_dict()
        except Exception as e:
            self.log_message('Настройки не скопировались', e)

        # ТермоЭДС
        for i in range(int(self.settings["n_cycles"])):  # Количество полных цилов термо эдс
            # Включаем релюшки
            # ! Добавить pause через try except
            self.toggle_relay("heater", "on")

            # Включаем нагреватель
            # ! Добавить pause через try except
            self.control_heater(channel=self.settings["ch_ip1"],
                                voltage=self.settings["u_ip1"],
                                state="on")

            # Основные измерения нагрева для термоЭДС
            for _ in range(int(self.settings["n_heat"])):
                self.termoemf_step()
                self.meas_number += 1
                self.number += 1
                self.app_instance.start_line_le.setText(str(self.number))

            # Выключаем нагреватель
            # ! Добавить pause через try except
            self.control_heater(channel=self.settings["ch_ip1"],
                                voltage="0",
                                state="on")

            # Выключаем релюшки
            # ! Добавить pause через try except
            self.toggle_relay("heater", "off")

            # Основные измерения охлаждения для термоЭДС
            for _ in range(int(self.settings["n_heat"])):
                self.termoemf_step()
                self.meas_number += 1
                self.number += 1
                self.app_instance.start_line_le.setText(str(self.number))

            time.sleep(int(self.app_instance.pause_s.text()))

        # Измерения сопротивления
        if self.app_instance.rele_cb.isChecked():
            # Включаем релюшки
            # ! Добавить pause через try except
            self.toggle_relay("sample", "on")

            # Включаем нагреватель
            # ! Добавить pause через try except
            self.control_heater(channel=self.settings["ch_ip2"],
                                voltage=self.settings["u_ip2"],
                                state="on")

            # Основные измерения сопротивления
            for _ in range(int(self.settings["n_r_updown"])):
                self.resistance_step()
                if not self.change_volt_flag:
                    # Включаем релюшки
                    # ! Добавить pause через try except
                    self.toggle_relay("current", "on")
                else:
                    # Выключаем релюшки
                    # ! Добавить pause через try except
                    self.toggle_relay("current", "off")

            # Выключаем нагреватель
            # ! Добавить pause через try except
            self.control_heater(channel=self.settings["ch_ip2"],
                                voltage="0",
                                state="on")
            # Выключаем релюшки
            # ! Добавить pause через try except
            self.toggle_relay("sample", "off")

        else:
            # Включаем нагреватель
            # ! Добавить pause через try except
            self.control_heater(channel=self.settings["ch_ip2"],
                                voltage=self.settings["u_ip2"],
                                state="on")

            for _ in range(int(self.settings["n_r_up"])):
                self.resistance_step()

            # Выключаем нагреватель
            # ! Добавить pause через try except
            self.control_heater(channel=self.settings["ch_ip2"],
                                voltage="0",
                                state="on")

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
        # Номер строки в Excel
        self.number = int(self.app_instance.start_line_le.text())

        # ! Номер строки эксперимента
        # self.meas_number  # Добавить в Excel

        # Время 1
        time1 = time.time() - self.app_instance.start_time

        # R термометра 1
        termometer1 = self.temperature()

        # Термопары
        try:
            all_tc = self.termoemf()
        except Exception as e:
            self.log_message("Ошибка termoemf", e)
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

        print(f"{self.number} | {self.meas_number} | {time1} | {termometer1[0]} | {r1[0]} | {r2[0]} | "
              f"{r3[0]} | {r4[0]} | {termometer2[0]} | {time2} | {system_time}")

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
        # Номер строки в Excel
        self.number = int(self.app_instance.start_line_le.text())
        # ! Номер строки эксперимента
        # self.meas_number  # Добавить в Excel

        # Время 1
        time1 = time.time() - self.app_instance.start_time

        # R термометра 1
        termometer1 = self.temperature()

        # Сопротивления
        try:
            all_res = self.resistance()
        except Exception as e:
            self.log_message("Ошибка resistance", e)
        r5 = [i for i in all_res["ch5"]]
        r6 = [i for i in all_res["ch6"]]

        # Катушка
        kat = int(self.app_instance.r_cell.text())

        # R термометра 2
        termometer2 = self.temperature()

        # Время 2
        time2 = time.time() - self.app_instance.start_time

        # Системное время
        system_time = str(time.strftime("%d.%m.%Y  %H:%M:%S", time.localtime()))

        print(f"{self.number} | {self.meas_number} | {time1} | {termometer1[0]} | {r5[0]} | {r6[0]} | "
              f"{kat} | {termometer2[0]} | {time2} | {system_time}")

    def temperature(self):
        """Функция измерения температуры"""
        if self.app_instance.combobox_scan.currentText() == "keithley2010":
            instrument = Keithley2010(self)
        elif self.app_instance.combobox_scan.currentText() == "Rigol":
            instrument = Rigol(self)
        else:
            self.fres_value = "Error"
            return self.fres_value
        try:
            instrument.set_fres_parameters(float(self.settings['nplc_term']),
                                           int(self.settings['ch_term1']),
                                           range=0,
                                           delay=0)
        except Exception as e:
            self.log_message("Ошибка задания параметров", e)
        try:
            self.fres_value = instrument.measure(1)
        except Exception as e:
            self.log_message("Ошибка измерения", e)
        instrument.reset()  # Сброс настроек перед напряжением

        return self.fres_value

    def resistance(self):
        """Функция измерения сопротивления"""
        res_results = {}
        if self.app_instance.combobox_scan.currentText() == "keithley2010":
            instrument = Keithley2010(self)
        elif self.app_instance.combobox_scan.currentText() == "Rigol":
            instrument = Rigol(self)
        else:
            res_results["ch5"] = "Error"
            return res_results

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
                    instrument.measure(
                        meas_count=(int(self.settings["n_read_ch56"]) - 1)))  # 4 оставшихся измерения
            else:
                continue

        instrument.reset()  # Сброс настроек перед сопротивлением

        return res_results

    def termoemf(self):
        """Функция измерения термоЭДС"""
        termoemf_results = {}
        if self.app_instance.combobox_scan.currentText() == "keithley2010":
            instrument = Keithley2010(self)
        elif self.app_instance.combobox_scan.currentText() == "Rigol":
            instrument = Rigol(self)
        else:
            termoemf_results["ch1"] = "Error"
            return termoemf_results

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
                        instrument.measure(
                            meas_count=(int(self.settings["n_read_ch12"]) - 1)))  # 4 оставшихся измерения
                else:
                    continue
            else:
                if int(self.settings["n_read_ch34"]) > 1:
                    instrument.set_dcv_parameters(float(nplc_line_edit),
                                                  int(ch_line_edit),
                                                  float(range_line_edit),
                                                  delay=0)  # Остальные измерения
                    termoemf_results[f"ch{i}"].extend(
                        instrument.measure(
                            meas_count=(int(self.settings["n_read_ch34"]) - 1)))  # 4 оставшихся измерения
                else:
                    continue

        instrument.reset()  # Сброс настроек перед сопротивлением
        # print("termoemf_results:" + termoemf_results)
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
