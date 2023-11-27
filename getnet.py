from vars import *

def lergetnet():
    global arquivogetnet
    arquivogetnet = filedialog.askopenfilename()
    label = Label(janela, text="{} CARREGADO".format(arquivogetnet), font="Arial 7")
    label.grid(column=0, row=2)

async def getnet():
    global arquivogetnet
    pasta = Book(arquivogetnet)
    planilha = pasta.sheets[1]

    execute_query(connection, 'USE fluxodecaixa;')

    lr = planilha.range('D4').end('down').row

    for i in range(5, lr):
        cell = planilha.range('B{}'.format(i)).value
        if cell == None:
            valor = planilha.range('D{}'.format(i)).value
            date = datetime.strptime(planilha.range('B{}'.format(i - 1)).value, '%d/%m/%Y').date()
            datagetnet.append(date)
            valorgetnet.append(valor)

    for i in range(0, len(datagetnet)):
        execute_query(connection, "INSERT INTO getnet (data, valor) VALUES ('{}', '{}');".format(datagetnet[i], valorgetnet[i]))

    pasta.close()
    os.system('taskkill /im Excel.exe')
    await asyncio.sleep(0)