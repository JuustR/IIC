"""
Основной блок программы, отвечающий за запуск всех элементов
"""

import PyQt6
import sys
import os
import win32com.client as win32
import pathlib
import time

from PyQt6.QtWidgets import QApplication
from GUI import App

if __name__ == "__main__":
    app = QApplication(sys.argv)
    ex = App()

    sys.exit(app.exec())