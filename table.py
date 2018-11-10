from openpyxl import Workbook


def create_table():
    wb = Workbook()
    ws = wb.active
    ws.title = 'nro serie'
    c_time = 'HORA'
    c_date = 'FECHA'
    c_alarm = 'ALARMA'
    c_activo = 'ESTADO'
    ws.cell(1, 1, c_date)
    ws.cell(1, 2, c_time)
    ws.cell(1, 3, c_alarm)
    ws.cell(1, 4, c_activo)

    return wb


def load_table(wb, fila, log):
    ws = wb.active
    for i in range(1, len(log) + 1):
        ws.cell(fila, i, log[i - 1])
    # wb.save('log.xlsx')
