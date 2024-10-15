"""
Чистовой класс для измерений
SettingsManager принимать из UI
"""

class Measurements:
    def __init__(self, instruments):
        self.instruments = instruments

    def apply_settings(self, settings_manager):
        """Применяем настройки ко всем приборам из settings_manager"""

        for instrument_name, instrument in self.instruments.items():
            settings = settings_manager.get_settings_for(instrument_name)

            if instrument:
                instrument.write(settings['reset'])
                instrument.write(settings['mode'])
                instrument.write(f"VOLT {settings['voltage']}")
                instrument.write(f"CURR {settings['current']}")

    def start_measurements(self):
        #! Переписать согласно experiment.py
        """Запускаем измерения для всех подключенных приборов"""

        for instrument_name, instrument in self.instruments.items():
            if instrument:
                voltage = instrument.query("MEAS:VOLT?")
                current = instrument.query("MEAS:CURR?")
                print(f"{instrument_name} - Напряжение: {voltage}, Ток: {current}")

class SettingsManager:
    def __init__(self):
        # Пример настроек для каждого прибора
        self.settings = {
            'keithley': {
                'reset': '*RST',
                'mode': ':SOUR:FUNC VOLT',
                'voltage': 5.0,
                'current': 0.1
            },
            'akip': {
                'reset': '*RST',
                'mode': 'VOLT',
                'voltage': 3.3,
                'current': 0.5
            },
            'e36312a': {
                'reset': '*RST',
                'mode': 'OUTP ON',
                'voltage': 12.0,
                'current': 1.0
            }
        }

    def get_settings_for(self, instrument_name):
        """Возвращаем настройки для указанного прибора"""
        return self.settings.get(instrument_name, None)