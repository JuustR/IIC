"""
Класс содержащий основной цикл измерений

"""

import time
import requests
import pyvisa
from PyQt6.QtCore import QObject

from PyQt6.QtCore import QThread, pyqtSignal

from Config.Keithley2010 import Keithley2010
from Config.Rigol import Rigol


class MeasurementThread(QThread):
    log_signal = pyqtSignal()
    finished_signal = pyqtSignal()

    def __init__(self, meas_instance):
        super().__init__()
        self.meas_instance = meas_instance
        self.running = True

    def run(self):
        try:
            while self.running:
                self.meas_instance.cycle_S_R()
        except Exception as e:
            self.log_signal.emit(f"Ошибка: {e}")
        finally:
            self.running = False
            self.meas_instance.pause()
            self.finished_signal.emit()

    def stop(self):
        """
        Завершает работу потока
        """
        self.running = False
        self.quit()
        self.wait()
        self.meas_instance.pause()


class Measurements(QObject):
    """
    Класс для работы с измерениями

    Note: Для связи с GUI использовать сигналы и слоты(не знаю что это), т.к. они не вызывают ошибки,
    связанные с многозадачностью/многопоточностью
    """
    update_excel_signal = pyqtSignal(int, int, object)  # Сигнал для обновления GUI
    update_values_signal = pyqtSignal(dict)  # Сигнал для обновления различных значений, напр. номера строки

    def __init__(self, app_instance):
        super().__init__()
        self.rm = pyvisa.ResourceManager()  # Инициализируем ResourceManager

        self.app_instance = app_instance
        self.log_signal = app_instance.log_signal
        self.settings = self.app_instance.settings_dict
        self.inst_list = self.app_instance.inst_list
        self.powersource_list = self.app_instance.powersource_list
        self.excel_cash = self.app_instance.excel_cash

        # Инициализация ИП
        if "AKIP" in self.powersource_list:
            self.AKIP = self.rm.open_resource(self.powersource_list["AKIP"])
        else:
            self.pause()

        if self.app_instance.combobox_scan.currentText() == "keithley2010":
            self.instrument = Keithley2010(self)
        elif self.app_instance.combobox_scan.currentText() in ["Rigol", "keysight"]:
            self.instrument = Rigol(self)
        else:
            return

        self.change_volt_flag = False  # Флаг отвечающий за переключение направления тока
        self.cash_flag = False

        self.fres_value = None
        self.number = self.app_instance.start_line_le.text()
        self.meas_number = 1

    def toggle_relay(self, relay_type, state):
        """
        Управление релюшками, где current - направление тока, sample - ток, heater - нагреватель
        """
        d = {
            "current": "current_switch",
            "sample": "sample_switch",
            "heater": "heater_current_switch",
        }
        try:
            action = d[relay_type]
            requests.get(f'http://192.168.0.101:10500/turn_{state}_{action}')
        except Exception as e:
            # self.pause()  #! Можно раскоментить, наверное
            self.log_message(f"Ошибка переключения реле {relay_type} ({state})", e)

    def control_heater(self, channel, voltage, state):
        """
        Управление нагревателем
        """
        try:
            if self.app_instance.combobox_power.currentText() == "AKIP":
                self.AKIP.write(f'INSTrument:NSELect {channel}')
                self.AKIP.write(f'APPL CH{channel},{voltage},1')
                self.AKIP.write(f'CHANnel:OUTPut {1 if state == "on" else 0}')
        except Exception as e:
            # self.pause()  #! Можно раскоментить, наверное
            self.log_message(f"Ошибка управления нагревателем ({state})", e)

    def log_message(self, message, exception=None):
        error_message = f"{message}"
        if exception:
            error_message += f"\n{exception}"
        self.log_signal.emit(error_message)

    def pause(self):
        """
        Функция паузы, останавливающая измерения и обнуляющая всё
        """
        self.app_instance.working_flag = False
        if "AKIP" in self.powersource_list:
            try:
                self.control_heater(channel=1, voltage=0, state="off")
                self.control_heater(channel=2, voltage=0, state="off")
                self.control_heater(channel=3, voltage=0, state="off")
                self.AKIP.write("*rst")  # На всякий случай)
            except Exception as e:
                self.log_message("Ошибка отключения нагревателя", e)
        if self.app_instance.rele_cb.isChecked():
            try:
                self.toggle_relay("heater", "off")
                self.toggle_relay("sample", "off")
                self.toggle_relay("current", "off")
            except Exception as e:
                self.log_message("Ошибка отключения релюшек\nВыключите их вручную для корректной работы", e)

    def cycle_S_R(self):
        """
        Основной цикл измерений термоЭДС и сопротивления
        """
        while self.app_instance.working_flag and self.app_instance.measurement_thread.running:
            # Проверка обновления настроек
            try:
                if self.app_instance.settings_changed_flag:
                    self.app_instance.copy_settings_to_dict()
            except Exception as e:
                self.log_message('Настройки не скопировались', e)

            try:
                if self.app_instance.startline_changed_flag:
                    self.number = self.app_instance.start_line_le.text()
            except Exception as e:
                self.log_message('Начальная строка не изменилась (наверное)', e)

            # ТермоЭДС
            for i in range(int(self.settings["n_cycles"])):  # Количество полных циклов термо эдс

                if not self.app_instance.measurement_thread.running or not self.app_instance.working_flag:
                    # self.log_message("Цикл измерений прерван на измерении термоЭДС")
                    break

                # Включаем релюшки
                # ! Если измерения без релюшек, то нужен try(наверное)
                self.toggle_relay("heater", "on")

                # Включаем нагреватель
                self.control_heater(channel=self.settings["ch_ip1"],
                                    voltage=self.settings["u_ip1"],
                                    state="on")

                # Основные измерения нагрева для термоЭДС
                for _ in range(int(self.settings["n_heat"])):
                    # Условия остановки измерений
                    if not self.app_instance.measurement_thread.running or not self.app_instance.working_flag:
                        # self.log_message("Цикл измерений прерван.")
                        break

                    self.termoemf_step()
                    self.meas_number += 1
                    self.number += 1
                    # self.update_values_signal.emit({"start_line_le": str(self.number)})
                    self.app_instance.start_line_le.setText(str(self.number))

                # Выключаем нагреватель
                self.control_heater(channel=self.settings["ch_ip1"],
                                    voltage="0",
                                    state="on")

                # Выключаем релюшки
                self.toggle_relay("heater", "off")

                # Основные измерения охлаждения для термоЭДС
                for _ in range(int(self.settings["n_cool"])):
                    # Условия остановки измерений
                    if not self.app_instance.measurement_thread.running or not self.app_instance.working_flag:
                        # self.log_message("Цикл измерений прерван на измерении термоЭДС")
                        break

                    self.termoemf_step()
                    self.meas_number += 1
                    self.number += 1
                    # self.update_values_signal.emit({"start_line_le": str(self.number)})
                    self.app_instance.start_line_le.setText(str(self.number))

                time.sleep(int(self.app_instance.pause_s.text()))

            # Измерения сопротивления
            # Условия остановки измерений
            if not self.app_instance.measurement_thread.running or not self.app_instance.working_flag:
                # self.log_message("Цикл измерений прерван после измерения термоЭДС")
                break

            if self.app_instance.rele_cb.isChecked():
                # Включаем релюшки
                self.toggle_relay("sample", "on")

                # Включаем нагреватель
                self.control_heater(channel=self.settings["ch_ip2"],
                                    voltage=self.settings["u_ip2"],
                                    state="on")

                # Основные измерения сопротивления
                for _ in range(int(self.settings["n_r_updown"]) * 2):
                    # Условия остановки измерений
                    if not self.app_instance.measurement_thread.running or not self.app_instance.working_flag:
                        # self.log_message("Цикл измерений прерван на сопротивлении")
                        break

                    self.resistance_step()
                    if not self.change_volt_flag:
                        # Включаем релюшки
                        self.toggle_relay("current", "on")
                        self.change_volt_flag = True
                    else:
                        # Выключаем релюшки
                        self.toggle_relay("current", "off")
                        self.change_volt_flag = False

                    self.meas_number += 1
                    self.number += 1
                    # self.update_values_signal.emit({"start_line_le": str(self.number)})
                    self.app_instance.start_line_le.setText(str(self.number))

                # Выключаем нагреватель
                self.control_heater(channel=self.settings["ch_ip2"],
                                    voltage="0",
                                    state="on")
                # Выключаем релюшки
                self.toggle_relay("sample", "off")

            else:
                # Включаем нагреватель
                self.control_heater(channel=self.settings["ch_ip2"],
                                    voltage=self.settings["u_ip2"],
                                    state="on")

                for _ in range(int(self.settings["n_r_up"])):
                    # Условия остановки измерений
                    if not self.app_instance.measurement_thread.running or not self.app_instance.working_flag:
                        # self.log_message("Цикл измерений прерван на сопротивлении")
                        break

                    self.resistance_step()

                    self.meas_number += 1
                    self.number += 1
                    # self.update_values_signal.emit({"start_line_le": str(self.number)})
                    self.app_instance.start_line_le.setText(str(self.number))

                # Выключаем нагреватель
                self.control_heater(channel=self.settings["ch_ip2"],
                                    voltage="0",
                                    state="on")

            # Условия остановки измерений
            if not self.app_instance.measurement_thread.running or not self.app_instance.working_flag:
                self.log_message("Цикл измерений прерван.")
                # ! Смещение на невыполненые строки (нужно тестить)
                step = (int(self.settings["n_cycles"]) * (int(self.settings["n_heat"]) + int(self.settings["n_cool"])) +
                        int(2 * self.settings["n_r_updown"] if self.app_instance.rele_cb.isChecked() else self.settings["n_r_up"]))
                if (self.number - 10) % step == 0:
                    # self.update_values_signal.emit({"start_line_le": str(self.number)})
                    self.app_instance.start_line_le.setText(str(self.number))
                else:
                    self.number += step - (self.number - 10) % step
                    # self.update_values_signal.emit({"start_line_le": str(self.number)})
                    self.app_instance.start_line_le.setText(str(self.number))
                break

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
        start_row = 1

        # Номер строки эксперимента
        self.update_excel_signal.emit(self.number, start_row, self.meas_number)
        start_row += 1

        # Время 1
        time1 = time.time() - self.app_instance.start_time
        self.update_excel_signal.emit(self.number, start_row, time1)
        start_row += 1

        # R термометра 1
        termometer1 = self.temperature()
        self.update_excel_signal.emit(self.number, start_row, termometer1)
        start_row += 1

        self.instrument.reset()

        # Термопары и между ними
        try:
            all_tc = self.termoemf()

            r1 = [i for i in all_tc["ch1"]]
            r2 = [i for i in all_tc["ch2"]]
            for i in range(len(r1)):
                self.update_excel_signal.emit(self.number, start_row, r1[i])
                start_row += 1
            for i in range(len(r2)):
                self.update_excel_signal.emit(self.number, start_row, r2[i])
                start_row += 1

            r3 = [i for i in all_tc["ch3"]]
            r4 = [i for i in all_tc["ch4"]]
            for i in range(len(r3)):
                self.update_excel_signal.emit(self.number, start_row, r3[i])
                start_row += 1
            for i in range(len(r4)):
                self.update_excel_signal.emit(self.number, start_row, r4[i])
                start_row += 1
        except Exception as e:
            self.log_message("Ошибка termoemf", e)

        self.instrument.reset()

        # Пропуск измерений сопротивления + катушка
        start_row += int(self.settings["n_read_ch56"]) * 2 + 1

        # R термометра 2
        termometer2 = self.temperature()
        self.update_excel_signal.emit(self.number, start_row, termometer2)
        start_row += 1

        # Время 2
        time2 = time.time() - self.app_instance.start_time
        self.update_excel_signal.emit(self.number, start_row, time2)
        start_row += 1

        # Системное время
        system_time = str(time.strftime("%d.%m.%Y  %H:%M:%S", time.localtime()))
        self.update_excel_signal.emit(self.number, start_row, system_time)

        # print(f"{self.number} | {self.meas_number} | {time1} | {termometer1[0]} | {r1[0]} | {r2[0]} | "
        #       f"{r3[0]} | {r4[0]} | {termometer2[0]} | {time2} | {system_time}")

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
        start_row = 1

        # Номер строки эксперимента
        self.update_excel_signal.emit(self.number, start_row, self.meas_number)
        start_row += 1

        # Время 1
        time1 = time.time() - self.app_instance.start_time
        self.update_excel_signal.emit(self.number, start_row, time1)
        start_row += 1

        # R термометра 1
        termometer1 = self.temperature()
        self.update_excel_signal.emit(self.number, start_row, termometer1)
        start_row += 1

        self.instrument.reset()

        # Пропуск измерений термоЭДС
        start_row += int(self.settings["n_read_ch12"]) * 2 + int(self.settings["n_read_ch34"]) * 2

        # Сопротивления
        try:
            all_res = self.resistance()
        except Exception as e:
            self.log_message("Ошибка resistance", e)
        r5 = [i for i in all_res["ch5"]]
        r6 = [i for i in all_res["ch6"]]
        for i in range(len(r5)):
            self.update_excel_signal.emit(self.number, start_row, r5[i])
            start_row += 1
        for i in range(len(r6)):
            self.update_excel_signal.emit(self.number, start_row, r6[i])
            start_row += 1

        self.instrument.reset()

        # Катушка
        kat = int(self.app_instance.r_cell.text())
        self.update_excel_signal.emit(self.number, start_row, kat)
        start_row += 1

        # R термометра 2
        termometer2 = self.temperature()
        self.update_excel_signal.emit(self.number, start_row, termometer2)
        start_row += 1

        # Время 2
        time2 = time.time() - self.app_instance.start_time
        self.update_excel_signal.emit(self.number, start_row, time2)
        start_row += 1

        # Системное время
        system_time = str(time.strftime("%d.%m.%Y  %H:%M:%S", time.localtime()))
        self.update_excel_signal.emit(self.number, start_row, system_time)

        # print(f"{self.number} | {self.meas_number} | {time1} | {termometer1[0]} | {r5[0]} | {r6[0]} | "
        #       f"{kat} | {termometer2[0]} | {time2} | {system_time}")

    def temperature(self):
        """
        Функция измерения температуры
        """
        try:
            self.instrument.set_fres_parameters(float(self.settings['nplc_term']),
                                           int(self.settings['ch_term1']),
                                           range=0,
                                           delay=0)
        except Exception as e:
            self.log_message("Ошибка задания параметров", e)
        try:
            self.fres_value = self.instrument.measure(1)
        except Exception as e:
            self.log_message("Ошибка измерения", e)

        self.instrument.trig_rigol()
        # self.instrument.reset()  # Сброс настроек перед напряжением

        return self.fres_value

    def resistance(self):
        """
        Функция измерения сопротивления
        """
        res_results = {}

        for i in range(5, 7):
            ch_line_edit = self.settings[f"ch{i}"]
            delay_line_edit = self.settings[f"delay_ch{i}"]
            range_line_edit = self.settings[f"range_ch{i}"]
            nplc_line_edit = self.settings[f"nplc_ch{i}"]

            self.instrument.set_dcv_parameters(float(nplc_line_edit),
                                          int(ch_line_edit),
                                          float(range_line_edit),
                                          float(delay_line_edit))  # Первая задержка
            res_results[f"ch{i}"] = self.instrument.measure(meas_count=1)  # Первое измерение

            # Измерения для больше чем одного read
            if int(self.settings["n_read_ch56"]) > 1:
                # self.instrument.set_dcv_parameters(float(nplc_line_edit),
                #                               int(ch_line_edit),
                #                               float(range_line_edit),
                #                               delay=0)  # Остальные измерения
                res_results[f"ch{i}"].extend(
                    self.instrument.measure(
                        meas_count=(int(self.settings["n_read_ch56"]) - 1)))
            else:
                continue

            self.instrument.trig_rigol()
        # self.instrument.reset()  # Сброс настроек перед сопротивлением

        return res_results

    def termoemf(self):
        """
        Функция измерения термоЭДС
        """
        termoemf_results = {}

        for i in range(1, 5):
            ch_line_edit = self.settings[f"ch{i}"]
            delay_line_edit = self.settings[f"delay_ch{i}"]
            range_line_edit = self.settings[f"range_ch{i}"]
            nplc_line_edit = self.settings[f"nplc_ch{i}"]

            self.instrument.set_dcv_parameters(float(nplc_line_edit),
                                          int(ch_line_edit),
                                          float(range_line_edit),
                                          float(delay_line_edit))  # Первая задержка
            termoemf_results[f"ch{i}"] = self.instrument.measure(meas_count=1)  # Первое измерение

            # Измерения для больше чем одного read
            if i < 3:
                if int(self.settings["n_read_ch12"]) > 1:
                    # self.instrument.set_dcv_parameters(float(nplc_line_edit),
                    #                               int(ch_line_edit),
                    #                               float(range_line_edit),
                    #                               delay=0)  # Остальные измерения
                    termoemf_results[f"ch{i}"].extend(
                        self.instrument.measure(
                            meas_count=(int(self.settings["n_read_ch12"]) - 1)))
                else:
                    continue
                # self.instrument.trig_rigol()
            else:
                if int(self.settings["n_read_ch34"]) > 1:
                    # self.instrument.set_dcv_parameters(float(nplc_line_edit),
                    #                               int(ch_line_edit),
                    #                               float(range_line_edit),
                    #                               delay=0)  # Остальные измерения
                    termoemf_results[f"ch{i}"].extend(
                        self.instrument.measure(
                            meas_count=(int(self.settings["n_read_ch34"]) - 1)))
                else:
                    continue

            self.instrument.trig_rigol()

        # self.instrument.reset()  # Сброс настроек перед сопротивлением

        return termoemf_results

    def add_dcv(self):
        pass

