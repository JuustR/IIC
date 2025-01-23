import win32com.client as win32
import os

class Chmo:
    def __init__(self):
        self.excel = None
        self.excel = win32.Dispatch('Excel.Application')
        self.excel.Visible = 1



class Hui:
    def __init__(self, app_instance: Chmo):
        self.app_instance = app_instance
        script_path = os.path.dirname(os.path.abspath(__file__))
        self.wb = self.app_instance.excel.Workbooks.Open(os.path.join(script_path, "..", "templates", "Base_temp.xltm"))
        self.ws = self.wb.Worksheets(1)
        print(self.ws)
        try:
            self.ws.Cells(1,1).Value = "Hui"
            print(1)
        except Exception as e:
            print(2)
            print(e)
        print(self.ws)

if __name__ == "__main__":
    x = Chmo()
    h = Hui(x)