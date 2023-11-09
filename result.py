import xlwings

from query import *

def result():
    data_atual = datetime.now()
    app = xlwings.App()
    pasta = app.books.add()
    planilha = pasta.sheets.active
    contdata = 3
    contdata2 = 1

    planilha.range('A1').value = "DATA"
    planilha.range('B1').value = "GETNET"
    planilha.range('C1').value = "SAFRAPAY"
    planilha.range('D1').value = "COBRANÇA BB"
    planilha.range('E1').value = "COBRANÇA SAFRA"
    planilha.range('F1').value = "COBRANÇA SANTANDER"
    planilha.range('G1').value = "DEPOSITOS"
    planilha.range('H1').value = "DESPESAS"

    planilha.range('A2').value = data_atual.strftime("%x")

    for i in range(0, 1460):
        incremento = data_atual + timedelta(days=contdata2)
        convertido = incremento.date()
        planilha.range('A{}'.format(contdata)).value = convertido.strftime("%x")
        contdata += 1
        contdata2 += 1

    lr = planilha.range('A2').end('down').row