"""
Plans:
Use previous design
Add simple reader
watch how writer works in old Ruslan's app
add few classes like in observer pattern
read about OOP

Future:
add logs
"""
import time
import sys
from PyQt6.uic import loadUi #!
from Config.Instruments import InstrumentConnection
from Measurings import Measurements, SettingsManager
from GUI import App

from PyQt6.QtWidgets import QApplication

if __name__ == '__main__':
    app = QApplication(sys.argv)
    ex = App()
    """
    # Подключаем приборы
    instrument_connection = InstrumentConnection(loadUi('assets/mainIIC.ui'))
    instrument_connection.Connect_all() # Подключаем все приборы
    instrument_connection.Instr_check() # Вывод IDN всех подключенных приборов
    # Получаем подключенные приборы
    instruments = {
        'keithley2001': instrument_connection.keithley2001,
        'akip': instrument_connection.AKIP,
        'e36312a': instrument_connection.E36312A
        # Добавьте сюда любые другие приборы, которые у вас есть
    }

    # Создаем менеджер для приборов
    instrument_manager = Measurements(instruments)

    # Создаем менеджер настроек (если нужно)
    settings_manager = SettingsManager()

    # Применяем настройки ко всем приборам
    instrument_manager.apply_settings(settings_manager)

    """

    sys.exit(app.exec())
