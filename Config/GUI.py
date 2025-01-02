"""
Окно основного графического интерфейса и его настройки

Tasks:
1) Добавить словарь со всеми настройками, чтоб в него заносились настройки и
   чтоб была возможность обновлять их между циклами измерений
2) Исправить ошибку -213 Keithley и разобрться с даком
3) data И base_data реализовать через settings
4) Пауза
4.1) Либо сделать счетчик действий, чтоб следующий старт начинал измерения с
     определенного действия. При этом, когда задаётся начало строки счетчик сбрасывался
4.2) Либо ждать окончания измерения, но тогда реализовать выход из длительной задержки
5) Проверить работу lstn тк он вроде не выводит измеряемые значения на прибор
6) Возможно, вынести измерения в отдельный класс
"""

import os
import win32com.client as win32
import time
import json
import pyvisa

from PyQt6.uic import loadUi
from PyQt6.QtWidgets import (QMainWindow, QDialog, QFileDialog, QLineEdit)
from PyQt6.QtCore import QSettings

from Config.ChooseExcelDialog import ChooseExcelDialog
from Config.Instruments import InstrumentConnection
from Config.Keithley2001 import Keithley2001
from Config.Rigol import Rigol

class App(QMainWindow):
    """GUI основной страницы программы"""

    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)

        try:
            # Относительно долго грузится, можно разделить на 2 потока и добавить анимацию включения программы
            self.excel = win32.Dispatch('Excel.Application')  # Создаем COM объект
            self.excel.Visible = False  # Excel invisible
        except Exception as e:
            print(f"Очистите gen_py\nМожно это сделать запустив программу gen_cache_clear.py\n{e}")

        # Путь к основной директории
        current_dir = os.path.dirname(os.path.abspath(__file__))
        # Загрузка ui, путем выхода в основную директорию
        loadUi(os.path.join(current_dir, '..', 'assets', 'mainIIC2.ui'), self)

        self.setWindowTitle('IIC Measuring Program')

        # Хранилище настроек
        self.settings = QSettings("lab425", "IIC")

        # Подключаем основные кнопки к соответсвующим функциям
        self.choose_button.clicked.connect(self.on_choose_excel_clicked)
        self.instruments_button.clicked.connect(self.on_instruments_clicked)
        self.create_button.clicked.connect(self.on_create_clicked)
        self.start_line_button.clicked.connect(self.on_start_line_clicked)
        self.start_button.clicked.connect(self.on_start_clicked)

        # Подключаем кнопки меню настроек
        self.save_settings_pb.clicked.connect(self.save_settings)

        # Time and console start settings
        formatted_time = time.strftime("%d-%m-%Y %H:%M:%S", time.localtime())
        self.ConsolePTE.setPlainText(formatted_time + "\n" +
                                     """
Здравствуйте! Для начала работы:
1. Выберите/создайте шаблон Excel
2. Настройте приборы и Excel
  2.1. Выберите каналы
  2.2. Настройте параметры
  2.3. Подключитель к приборам
3.(необ.) Задайте начальную строчку Excel
4. Запускайте измерения
        """)

        # data - changeable, base_data - unchangeable REMEMBER IT!!!
        self.data = {"TempName": "Нет шаблона",
                     "MacrosName": "Нет макроса",
                     "FileName": time.strftime("%d-%m-%Y_Example", time.localtime())}
        self.base_data = {"TempName": "Нет шаблона",
                          "MacrosName": "Нет макроса",
                          "FileName": time.strftime("%d-%m-%Y_Example", time.localtime())}

        # Flag for start button
        self.working_flag = False
        self.data_reset_flag = False

        self.inst_dict = None

        # Загрузка настроек для Seebeck+R
        self.load_tab1_settings()

        self.show()

    def on_instruments_clicked(self):
        """Подключается ко всем доступным приборам, которые обнаружит"""
        #! Добавить ресет подключенных приборов
        ic = InstrumentConnection(self)
        self.inst_dict = ic.connect_all()
        connected_instruments = ', '.join(str(i) for i in self.inst_dict)
        self.ConsolePTE.appendPlainText(
            time.strftime("%H:%M:%S | ", time.localtime()) +
            'Подключенные приборы: ' + connected_instruments + "\n")

    def on_create_clicked(self):
        """Создаёт шаблон эксперимента"""
        self.ConsolePTE.appendPlainText(
            time.strftime("%H:%M:%S | ", time.localtime()) + 'НЕДОСТУПНО! Находится в разработке \n')
        pass

    def on_start_line_clicked(self):
        """Задаёт начало строки, по дефолту выставляет начало строки на 11 (реализовать по созданию Excel)"""
        pass

    def on_start_clicked(self):
        """Запускает эксперимент"""
        if not self.working_flag:
            self.working_flag = True
            self.start_button.setText('Стоп')

            # try расписать внутри функции

            try:
                # self.rigol_measurements()
                self.keithley_measurements()
            except Exception as e:
                print(e)

        else:
            self.working_flag = False
            self.start_button.setText('Старт')



        # self.start_fuct()

    def on_choose_excel_clicked(self):
        """Вызов диалога с созданием нового Excel"""
        self.ConsolePTE.appendPlainText(
            time.strftime("%H:%M:%S | ", time.localtime()) + 'Вызвали создание нового Excel\n')
        dlg = ChooseExcelDialog(self)
        dlg.exec()

    def save_settings(self):
        """По кнопке сохраняет настройки программы для Seebeck+R"""
        self.settings.setValue("combobox_scan", self.combobox_scan.currentText())
        self.settings.setValue("combobox_power", self.combobox_power.currentText())
        self.settings.setValue("rele_cb", self.rele_cb.isChecked())

        for i in range(1, 7):  #Ch1 - Ch6
            self.settings.setValue(f"ch{i}", self.findChild(QLineEdit, f"ch{i}").text())
            self.settings.setValue(f"delay_ch{i}", self.findChild(QLineEdit, f"delay_ch{i}").text())
            self.settings.setValue(f"range_ch{i}", self.findChild(QLineEdit, f"range_ch{i}").text())
            self.settings.setValue(f"nplc_ch{i}", self.findChild(QLineEdit, f"nplc_ch{i}").text())

        self.settings.setValue("n_read_ch12", self.n_read_ch12.text())
        self.settings.setValue("n_read_ch34", self.n_read_ch34.text())
        self.settings.setValue("n_read_ch56", self.n_read_ch56.text())
        self.settings.setValue("ch_term1", self.ch_term1.text())
        self.settings.setValue("ch_term2", self.ch_term2.text())
        self.settings.setValue("delay_term", self.delay_term.text())
        self.settings.setValue("range_term", self.range_term.text())
        self.settings.setValue("nplc_term", self.nplc_term.text())
        self.settings.setValue("ch_ip1", self.ch_ip1.text())
        self.settings.setValue("ch_ip2", self.ch_ip2.text())
        self.settings.setValue("u_ip1", self.u_ip1.text())
        self.settings.setValue("u_ip2", self.u_ip2.text())
        self.settings.setValue("n_cycles", self.n_cycles.text())
        self.settings.setValue("n_heat", self.n_heat.text())
        self.settings.setValue("n_cool", self.n_cool.text())
        self.settings.setValue("pause_s", self.pause_s.text())
        self.settings.setValue("pause_r", self.pause_r.text())
        self.settings.setValue("n_r_up", self.n_r_up.text())
        self.settings.setValue("n_r_updown", self.n_r_updown.text())
        self.settings.setValue("ip_rigol", self.ip_rigol.text())

        #Добавление текста в консоль
        self.ConsolePTE.appendPlainText(
            time.strftime("%H:%M:%S | ", time.localtime()) + 'Сохранили настройки программы\n')

        #! Добавить сравнение изменений и вывод их в Excel, например
        #!

    def load_tab1_settings(self):
        """Загружает сохранённые значения виджетов пока только для Seebeck+R"""
        self.combobox_scan.setCurrentText(self.settings.value("combobox_scan", ""))
        self.combobox_power.setCurrentText(self.settings.value("combobox_power", ""))
        self.rele_cb.setChecked(self.settings.value("rele_cb", "false") == "true")

        for i in range(1, 7):
            self.findChild(QLineEdit, f"ch{i}").setText(self.settings.value(f"ch{i}", ""))
            self.findChild(QLineEdit, f"delay_ch{i}").setText(self.settings.value(f"delay_ch{i}", ""))
            self.findChild(QLineEdit, f"range_ch{i}").setText(self.settings.value(f"range_ch{i}", ""))
            self.findChild(QLineEdit, f"nplc_ch{i}").setText(self.settings.value(f"nplc_ch{i}", ""))

        self.n_read_ch12.setText(self.settings.value("n_read_ch12", ""))
        self.n_read_ch34.setText(self.settings.value("n_read_ch34", ""))
        self.n_read_ch56.setText(self.settings.value("n_read_ch56", ""))
        self.ch_term1.setText(self.settings.value("ch_term1", ""))
        self.ch_term2.setText(self.settings.value("ch_term2", ""))
        self.ch_ip1.setText(self.settings.value("ch_ip1", ""))
        self.ch_ip2.setText(self.settings.value("ch_ip2", ""))
        self.u_ip1.setText(self.settings.value("u_ip1", ""))
        self.u_ip2.setText(self.settings.value("u_ip2", ""))
        self.delay_term.setText(self.settings.value("delay_term", ""))
        self.range_term.setText(self.settings.value("range_term", ""))
        self.nplc_term.setText(self.settings.value("nplc_term", ""))
        self.n_cycles.setText(self.settings.value("n_cycles", ""))
        self.n_heat.setText(self.settings.value("n_heat", ""))
        self.n_cool.setText(self.settings.value("n_cool", ""))
        self.pause_s.setText(self.settings.value("pause_s", ""))
        self.pause_r.setText(self.settings.value("pause_r", ""))
        self.n_r_up.setText(self.settings.value("n_r_up", ""))
        self.n_r_up.setText(self.settings.value("n_r_up", ""))
        self.ip_rigol.setText(self.settings.value("ip_rigol", ""))


    def keithley_measurements(self):
        # ! Добавить начало измерений в консоль
        # ! Номер строки, и время цифрами добавить
        # ! Изменить ренджи для каждой
        # !! Сделать вывод в Excel и сохранение в памяти python для кэша

        # Запись значений с прибора FRES - 6xDCV - FRES
        instr = Keithley2001(self)
        # ! Добавить range and delay
        instr.set_fres_parameters(float(self.nplc_term.text()),
                                  int(self.ch_term1.text()),
                                  range=0,
                                  delay=0)
        fres_res_1 = instr.measure(1)
        print(f"FRES on channel 101: {fres_res_1}")

        dcv_results = {}
        for i in range(1, 7):
            ch_line_edit = self.findChild(QLineEdit, f"ch{i}")  # Поиск элемента с именем ch{i}
            delay_line_edit = self.findChild(QLineEdit, f"dealy_ch{i}")
            range_line_edit = self.findChild(QLineEdit, f"range_ch{i}")
            nplc_line_edit = self.findChild(QLineEdit, f"nplc_ch{i}")

            instr.set_dcv_parameters(float(nplc_line_edit.text()),
                                 int(ch_line_edit.text()),
                                 float(range_line_edit.text()),
                                 float(delay_line_edit.text()))  # Первая задержка
            dcv_results[f"ch{i}"] = instr.measure(meas_count=1)  # Первое измерение

            # Измерения для больше чем одного read
            if i < 3:
                if int(self.n_read_ch12.text()) > 1:
                    instr.set_dcv_parameters(float(nplc_line_edit.text()),
                                         int(ch_line_edit.text()),
                                         float(range_line_edit.text()),
                                         delay=0)  # Остальные измерения
                    dcv_results[f"ch{i}"].extend(
                        instr.measure(meas_count=(int(self.n_read_ch12.text()) - 1)))  # 4 оставшихся измерения
                else:
                    continue
            elif 2 < i < 5:
                if int(self.n_read_ch34.text()) > 1:
                    instr.set_dcv_parameters(float(nplc_line_edit.text()),
                                         int(ch_line_edit.text()),
                                         float(range_line_edit.text()),
                                         delay=0)  # Остальные измерения
                    dcv_results[f"ch{i}"].extend(
                        instr.measure(meas_count=(int(self.n_read_ch34.text()) - 1)))  # 4 оставшихся измерения
                else:
                    continue
            else:
                if int(self.n_read_ch56.text()) > 1:
                    instr.set_dcv_parameters(float(nplc_line_edit.text()),
                                         int(ch_line_edit.text()),
                                         float(range_line_edit.text()),
                                         delay=0)  # Остальные измерения
                    dcv_results[f"ch{i}"].extend(
                        instr.measure(meas_count=(int(self.n_read_ch56.text()) - 1)))  # 4 оставшихся измерения
                else:
                    continue

        for channel, results in dcv_results.items():
            print(f"DCV on channel {channel}: {results}")

        instr.set_fres_parameters(float(self.nplc_term.text()), int(self.ch_term1.text()),
                                 range=0, delay=0)
        fres_result_2 = instr.measure(1)
        print(f"FRES on channel 101 (repeat): {fres_result_2}")

    def rigol_measurements(self):
        instr = Rigol(self)

        # Настройка и измерение 4-проводного сопротивления на канале 101
        instr.set_fres_parameters(
            float(self.nplc_term.text()),
            int(self.ch_term1.text()),
            range=0,
            delay=0
        )
        fres_res_1 = instr.measure(1) 
        print(f"FRES on channel 101: {fres_res_1}")

        # Словарь для хранения результатов измерений
        dcv_results = {}

        # Перебор каналов с параметрами
        for i in range(1, 7):
            ch_line_edit = self.findChild(QLineEdit, f"ch{i}")  # Поиск элемента с именем ch{i}
            delay_line_edit = self.findChild(QLineEdit, f"dealy_ch{i}")
            range_line_edit = self.findChild(QLineEdit, f"range_ch{i}")
            nplc_line_edit = self.findChild(QLineEdit, f"nplc_ch{i}")


            try:
                # Открытие канала
                instr.open_channel(ch_line_edit)

                # Настройка и измерение на текущем канале
                instr.set_dcv_parameters(
                    float(nplc_line_edit.text()),
                    int(ch_line_edit.text()),
                    float(range_line_edit.text()),
                    float(delay_line_edit.text())
                )
                dcv_results[f"ch{i}"] = instr.measure(meas_count=1)

                # Дополнительные измерения (если требуется)
                if i < 3:
                    additional_reads = int(self.n_read_ch12.text()) - 1
                elif 2 < i < 5:
                    additional_reads = int(self.n_read_ch34.text()) - 1
                else:
                    additional_reads = int(self.n_read_ch56.text()) - 1

                if additional_reads > 0:
                    instr.set_dcv_parameters(
                        float(nplc_line_edit.text()),
                        int(ch_line_edit.text()),
                        float(range_line_edit.text()),
                        delay=0  # Установка задержки для оставшихся измерений
                    )
                    dcv_results[f"ch{i}"].extend(instr.measure(meas_count=additional_reads))
            finally:
                # Закрытие канала
                instr.close_channel(ch_line_edit)

        # Вывод результатов измерений
        for channel, results in dcv_results.items():
            print(f"DCV on channel {channel}: {results}")

        # Повторное измерение 4-проводного сопротивления на канале 101
        instr.set_fres_parameters(
            float(self.nplc_term.text()),
            int(self.ch_term1.text()),
            range=0,
            delay=0
        )
        fres_result_2 = instr.measure(1)
        print(f"FRES on channel 101 (repeat): {fres_result_2}")
