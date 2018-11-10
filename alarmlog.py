import table


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
    is_hs = is_date = is_alarm = is_estado = True
    cad = []
    fecha = hora = alarm = estado = ''
    for i in line:
        if i not in ('_', '>', '[', ']'):
            if is_date:
                cad.append(i)
            elif is_hs:
                cad.append(i)
            elif is_alarm:
                cad.append(i)
            elif is_estado:
                cad.append(i)
        else:
            if i == '_':
                del cad[-1]
                fecha = ''.join(cad)
                cad = []
            elif i == '>':
                del cad[0]
                del cad[-1]
                hora = ''.join(cad)
                cad = []
            elif i == '[':
                del cad[0]
                del cad[-1]
                alarm = ''.join(cad)
                cad = []
            elif i == ']':
                estado = ''.join(cad)
                cad = []

    return [fecha, hora, alarm, estado]


def open_xlsm(fd):
    k = 1
    wb = table.create_table()
    m = open(fd)
    while True:
        k += 1
        line = m.readline()
        if line == '':
            break
        if line[-1] == '\n':
            line = line[:-1]
        print(line)
        alarm = disarm(line)
        table.load_table(wb, k, alarm)
    m.close()
    wb.save('log.xlsx')


def main():
    fd = 'log2.txt'
    open_xlsm(fd)

# log = '2018-07-10 _ 23:07:05 > AL: VOLUMEN MINUTO MÁXIMO [Activada]'
# alarm = disarm(log)
# print(alarm)
# ws = table.create_table()
# table.load_table(ws, 2, alarm)


if __name__ == '__main__':
    main()
