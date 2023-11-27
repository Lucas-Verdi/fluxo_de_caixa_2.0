from getnet import *


def lersafrapay():
    global arquivosafrapay
    arquivosafrapay = filedialog.askopenfilename()
    label = Label(janela, text="{} CARREGADO".format(arquivosafrapay), font="Arial 7")
    label.grid(column=0, row=4)


async def safrapay():
    global arquivosafrapay
    pasta = Book(arquivosafrapay)
    planilha = pasta.sheets[0]

    execute_query(connection, 'USE fluxodecaixa;')

    lr = planilha.range('F2').end('down').row

    soma = 0
    for i in range(2, lr + 1):
        data = planilha.range(f'F{i}').value
        data1 = planilha.range(f'F{i + 1}').value
        valor = planilha.range(f'L{i}').value
        soma += valor
        if data != data1 or data1 == None:
            databarra = data.replace('.', '/')
            date = datetime.strptime(databarra, '%d/%m/%Y').date()
            datasafrapay.append(date)
            valorsafrapay.append(soma)
            soma = 0

    for i in range(0, len(datasafrapay)):
        execute_query(connection, "INSERT INTO safra (data, valor) VALUES ('{}', '{}');".format(datasafrapay[i], valorsafrapay[i]))

    pasta.close()
    os.system('taskkill /im Excel.exe')
    await asyncio.sleep(0)