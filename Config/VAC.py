"""

"""

import os
import pyvisa
import time

from PyQt6.uic import loadUi
from PyQt6.QtWidgets import (QDialog, QFileDialog)
from Config.Keithley2010 import Keithley2010
from Config.Rigol import Rigol


class VAC(QDialog):
    def __init__(self, app_instance, parent=None):
        super().__init__(parent)
        self.app_instance = app_instance
        self.rm = pyvisa.ResourceManager()

        current_dir = os.path.dirname(os.path.abspath(__file__))
        loadUi(os.path.join(current_dir, '..', 'assets', 'chooseDialog.ui'), self)

        if self.app_instance.combobox_scan.currentText() == "keithley2010":
            self.instrument = Keithley2010(self)
        elif self.app_instance.combobox_scan.currentText() in ["Rigol", "keysight"]:
            self.instrument = Rigol(self)
        else:
            return

