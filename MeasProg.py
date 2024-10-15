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

from Instruments import InstrumentConnection
from Measurings import Measurements, SettingsManager
from GUI import App

from PyQt6.QtWidgets import QApplication

if __name__ == '__main__':
    app = QApplication(sys.argv)
    """
    # Подключаем приборы
    instrument_connection = InstrumentConnection()
    instrument_connection.connect_all()  # Подключаем все приборы
    # Получаем подключенные приборы
    instruments = {
        'keithley': instrument_connection.keithley,
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
    ex = App()
    sys.exit(app.exec())
