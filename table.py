from openpyxl import Workbook
from openpyxl.styles import PatternFill
from openpyxl.styles import Font
from openpyxl.styles import Alignment

campos = ['HORA', 'FECHA', 'CATEGORIA', 'ALARMA', 'ESTADO']
col = ['A1', 'B1', 'C1', 'D1', 'E1']


def create_table():
    wb = Workbook()
    ws = wb.active
    ws.title = 'nro serie'
    for i in range(len(campos)):
        ws.cell(1, i + 1, campos[i])
        ws[col[i]].fill = PatternFill(fgColor="0070C0", fill_type="solid")
        ws[col[i]].font = Font(color="FFFFFF")
        ws[col[i]].alignment = Alignment(horizontal='center')
    return wb


def load_table(wb, fila, log):
    ws = wb.active
    for i in range(1, len(log) + 1):
        ws.cell(fila, i, log[i - 1])
        # ws[col[i - 1]].fill = PatternFill(fgColor="0070C0", fill_type="solid")
        #if (i % 2) == 0:
            #pass
