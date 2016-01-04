from openpyxl import load_workbook

ruta = r'c:\users\rmasgo\desktop\Fact Deyre 2016.xlsx'
resultado = r'c:\users\rmasgo\desktop\Fact Deyre 2016 Resultado.xlsx'
#hoja = 'Hoja1'


def Ultimo(sheet, columna):  # Busca la ultima fila con datos
    a = 1
    col = columna + str(a)
    while sheet[col].value != None:
        a = a + 1
        col = columna + str(a)
    return a - 1


def Encontrar(sheet, columna, valor):  # Busca la ultima fila con datos
    a = 1
    col = columna + str(a)
    while sheet[col].value != None and sheet[col].value.strip().replace(",", "") != valor.strip().replace(",", ""):
        a = a + 1
        col = columna + str(a)
    if sheet[col].value == None: a = 0
    return a


def Comparar(sheet, columnaInicial = 'A',columnaResultado = 'C',columnaBuscar = 'E',columnaImporte = 'F'):
    i = 2
    col = columnaInicial + str(i)
    while sheet[col].value != None:
        valor = sheet[col].value
        col = columnaResultado + str(i)
        # Buscalos el valor en la columna E
        encontrado = Encontrar(sheet, columnaBuscar, valor)
        if encontrado == 0:  # Si no lo encontramos lo marcamos
            sheet[col] = 'X'
        else:
            colImporte = columnaImporte + str(encontrado)
            sheet[col].value = sheet[colImporte].value
        i = i + 1
        col = columnaInicial + str(i)
        print(col + ' ')

# Abrimos el excel
wb = load_workbook(filename=ruta)
# Obtenemos la hoja
#sheet = wb[hoja]
sheet = wb.active

print('Columna 1')
Comparar(sheet, 'A', 'C', 'E', 'F')
print('Columna 2')
Comparar(sheet, 'E', 'G', 'A', 'B')

wb.save(resultado)


print("-- FIN --")