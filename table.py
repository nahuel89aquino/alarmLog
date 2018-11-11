from openpyxl import Workbook
from openpyxl.styles import PatternFill
from openpyxl.styles import Font
from openpyxl.styles import Alignment

campos = ['HORA', 'FECHA', 'CATEGORIA', 'ALARMA', 'ESTADO']
col = ['A', 'B', 'C', 'D', 'E']


def create_table():
    wb = Workbook()
    ws = wb.active
    ws.title = 'nro serie'
    for i in range(len(campos)):
        celda = col[i] + '1'
        ws.cell(1, i + 1, campos[i])
        ws[celda].fill = PatternFill(fgColor="0070C0", fill_type="solid")
        ws[celda].font = Font(color="FFFFFF")
        ws[celda].alignment = Alignment(horizontal='center')
    return wb


def load_table(wb, fila, log):
    ws = wb.active
    for i in range(1, len(log) + 1):
        celda = col[i - 1] + str(fila)
        ws.cell(fila, i, log[i - 1])
        # print(col[i - 1])
        if (fila % 2) == 0:
            ws[celda].fill = PatternFill(fgColor="EEECE1", fill_type="solid")
        else:
            ws[celda].fill = PatternFill(fgColor="BFBFBF", fill_type="solid")

