from moduls import table
import os

class Fecha:
    def __init__(self, dia, hora, mes):
        self.dia = dia
        self.hora = hora
        self.mes = mes


class Hora:
    def __init__(self, hora, min, seg):
        self.hora = hora
        self.min = min
        self.seg = seg


def disarm(line):
    is_hs = is_date = is_alarm = True
    is_cat = False
    cad = []
    fecha = hora = alarm = estado = categoria = ''
    for i in line:
        if i not in ('_', '>', ':', '[', ']'):
            cad.append(i)
        else:
            if i == '_'and is_date:
                fecha = ''.join(cad)
                cad = []
                is_date = False
            elif i == '>' and is_hs:
                hora = ''.join(cad)
                cad = []
                is_hs = False
                is_cat = True
            elif i == ':' and is_cat:
                categoria = ''.join(cad)
                cad = []
                is_cat = False
            elif i == '[' and is_alarm:
                alarm = ''.join(cad)
                cad = []
                is_alarm = False
            elif i == ']':
                estado = ''.join(cad)
                cad = []
            else:
                cad.append(i)
    if estado == '':
        alarm = ''.join(cad)

    return [fecha, hora, categoria, alarm, estado]


def open_xlsx(fd, newdir,name):
    k = 1
    wb = table.create_table(name)
    m = fd
    while True:
        line = m.readline()
        if line == '':
            break
        if line[-1] == '\n'and line is not '\n':
            line = line[:-1]
            print(line)
            k += 1
            alarm = disarm(line)
            table.load_table(wb, k, alarm)
        m.flush()
    os.chdir(newdir)
    wb.save(name+'.xlsx')



