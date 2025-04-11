import sys
import random
import time
import os
import xlwings as xw
from PyQt6.QtWidgets import QApplication, QWidget, QPushButton, QVBoxLayout
from PyQt6.QtCore import QThread, pyqtSignal




class ExcelWriterThread(QThread):
    stop_signal = pyqtSignal()

    def __init__(self):
        super().__init__()
        self.running = True

    def run(self):
        while self.running:
            self.write_random_numbers()
            time.sleep(1)  # Пауза между записями

    def stop(self):
        self.running = False
        self.stop_signal.emit()

    def write_random_numbers(self):
        try:

            xw_wb = xw.Book(EXCEL_FILE) if os.path.exists(EXCEL_FILE) else xw.books.add()
            xw_ws = xw_wb.sheets[0]

            last_row = xw_ws.range("A" + str(xw_ws.cells.last_cell.row)).end("up").row + 1

            xw_ws.range((last_row, 1)).value = random.randint(1, 5)
            last_row += 1

            # xw_wb.save(EXCEL_FILE)

        except Exception as e:
            print("Ошибка работы с xlwings:", e)


class App(QWidget):
    def __init__(self):
        super().__init__()
        self.initUI()
        self.thread = None


    def initUI(self):
        layout = QVBoxLayout()

        self.start_btn = QPushButton("Старт")
        self.start_btn.clicked.connect(self.start_writing)
        layout.addWidget(self.start_btn)

        self.stop_btn = QPushButton("Стоп")
        self.stop_btn.clicked.connect(self.stop_writing)
        layout.addWidget(self.stop_btn)

        self.setLayout(layout)
        self.setWindowTitle("Excel Writer")
        self.resize(200, 100)

    def start_writing(self):
        if self.thread is None or not self.thread.isRunning():
            self.thread = ExcelWriterThread()
            self.thread.start()

    def stop_writing(self):
        if self.thread and self.thread.isRunning():
            self.thread.stop()
            self.thread = None


if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = App()
    window.show()
    sys.exit(app.exec())
