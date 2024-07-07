"""
New program for INH experiments
Don't forget about VAX!!!
"""
import time
import sys

# from Voltmeters import KeysightVoltmeter, KeithleyVoltmeter, RigolVoltmeter
# from PowerSupply import KeysightPowerSupply
from GUI import App
# from Excel import Excel

from PyQt6.QtWidgets import QApplication

if __name__ == '__main__':
    app = QApplication(sys.argv)
    ex = App()
    sys.exit(app.exec())
