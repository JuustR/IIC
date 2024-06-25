
import win32com.client as win32
import os


class Excel:
    def __init__(self):
        self.cash_row = []
        self.cash_column = []
        self.cash_value = []
        self.cash_flag = False
        #self.excel = win32.gencache.EnsureDispatch('Excel.Application')  # Создаем COM объект
        self.excel = win32.Dispatch('Excel.Application')  # Создаем COM объект
        self.excel.Visible = True
        """
        try:
            self.excel = win32.Dispatch('Excel.Application')  # Создаем COM объект
        except:
            self.excel = win32.gencache.EnsureDispatch('Excel.Application')  # Создаем COM объект
        """

    def open_temp(self, data):
        file_name = data[3]
        # self.excel.Visible = True
        wb = self.excel.Workbooks.Open(file_name)
        self.ws = wb.Worksheets(1)


    def create_excel(self, data):

        script_path = os.path.abspath(__file__)
        file_path = os.path.dirname(script_path) + "\\Measurements\\" + data[1] # without '.xlsm'

        if data[4] == None:
            workbook = self.excel.Workbooks.Open(os.path.dirname(script_path) + "\\templates\\Base_temp.xltm")
        else:
            workbook = self.excel.Workbooks.Open(data[4])
            # workbook.SaveAs(file_path, 52)

        # Формат .xlsm будет при 52, а .xlsx при 51
        if data[2]:
            try:
                vbacomponent = workbook.VBProject.VBComponents.Add(1)  # 1 = vbext_ct_StdModule
                vbacomponent.CodeModule.AddFromFile("C:\\Dima\\INH\\VBAcode.txt")
                print("Макросы успешно добавлены")
            except:
                print("Макросы не добавлены")
                print("В Центре управления безопасностью поставить галочку на Доверять доступ к объектной модели проектов VBA")
        else:
            print("Макросы не добавлены")
        workbook.SaveAs(file_path, 52)

    def calculation(self, val: bool):
        self.ws.EnableCalculation = val

    def write_cash(self):
        for i in range(0, len(self.cash_row)):
            self.ws.Cells(self.cash_row[i], self.cash_column[i]).Value = self.cash_value[i]
        self.cash_row.clear()
        self.cash_column.clear()
        self.cash_value.clear()
        self.cash_flag = False

    def write_value(self, row, column, value):
        #print(f'wtite_value row{row}, value:{value}')
        try:
            self.ws.Cells(row, column).Value = value
            if self.cash_flag:
                self.write_cash()
        except Exception as ex:
            print(ex)
            self.cash_flag = True
            self.cash_row.append(row)
            self.cash_column.append(column)
            self.cash_value.append(value)

    def create_formula(self, row, column, formula):
        # example formula "=SUM(A2:D4)"
        self.ws.Cells(row, column).Formula = formula

    def create_chart(self, range, title):
        self.ws.Shapes.AddChart().Select()
        # ChartType https://stackoverflow.com/questions/38813522/themed-chart-styles-for-excel-charts-using-vba
        self.excel.ActiveChart.ChartType = win32.constants.xlXYScatterLines
        self.excel.ActiveChart.SetSourceData(Source=self.ws.Range("A1:A5,D1:D5"))
        self.excel.ActiveChart.HasLegend = False
        self.excel.ActiveChart.HasTitle = True
        self.excel.ActiveChart.ChartTitle.Text = title

    def converter_number_to_letter(self, number):
        letters = ('A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M', 'N', 'O', 'P', 'Q', 'R', 'S',
                        'T', 'U', 'V', 'W', 'X', 'Y', 'Z', 'AA', 'AB', 'AC', 'AD', 'AE', 'AF', 'AG', 'AH', 'AI', 'AJ',
                        'AK', 'AL', 'AM', 'AN', 'AO', 'AP', 'AQ', 'AR', 'AS', 'AT', 'AU', 'AV', 'AW', 'AX', 'AY', 'AZ',
                        'BA', 'BB', 'BC', 'BD', 'BE', 'BF', 'BG', 'BH', 'BI', 'BJ', 'BK', 'BL', 'BM', 'BN', 'BO', 'BP',
                        'BQ', 'BR', 'BS', 'BT', 'BU', 'BV', 'BW', 'BX', 'BY', 'BZ', 'CA', 'CB', 'CC', 'CD', 'CE', 'CF',
                        'CG', 'CH', 'CI', 'CJ', 'CK', 'CL', 'CM', 'CN', 'CO', 'CP', 'CQ', 'CR', 'CS', 'CT', 'CU', 'CV',
                        'CW', 'CX', 'CY', 'CZ', 'DA', 'DB', 'DC' , 'DD', 'DE', 'DF', 'DG', 'DH', 'DI', 'DJ', 'DK', 'DL',
                        'DM', 'DN', 'DO', 'DP', 'DQ', 'DR', 'DS', 'DT', 'DU', 'DV', 'DW', 'DX', 'DY', 'DZ', 'EA', 'EB')
        return letters[number - 1]