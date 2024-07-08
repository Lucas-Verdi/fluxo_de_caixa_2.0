from vars import *

def lergetnet():
    global arquivogetnet
    arquivogetnet = filedialog.askopenfilename()
    label = Label(janela, text="{} CARREGADO".format(arquivogetnet), font="Arial 7")
    label.grid(column=0, row=2)

async def getnet():
    global arquivogetnet
    if arquivogetnet == None:
        print('Vazio')
    else:
        pasta = Book(arquivogetnet)
        planilha = pasta.sheets[2]
        execute_query(connection, 'USE fluxodecaixa;')

        lr = planilha.range('A9').end('down').row
        acumulo = 0

        celula = planilha.range('A9').value
        if celula == None:
            print('VAZIO')
        else:
            for i in range(9, lr):
                cell = planilha.range('D{}'.format(i)).value
                if cell == None:
                    data_temporaria = planilha.range('D{}'.format(i + 1)).value
                    data_temporaria_2 = planilha.range('D{}'.format(i - 1)).value
                    if data_temporaria == data_temporaria_2:
                        valor = planilha.range('K{}'.format(i)).value
                        acumulo += valor
                    elif data_temporaria != data_temporaria_2:
                        valor = planilha.range('K{}'.format(i)).value
                        acumulo += valor
                        date = planilha.range('D{}'.format(i - 1)).value
                        datagetnet.append(date)
                        valorgetnet.append(acumulo)
                        acumulo = 0
                

        for i in range(0, len(datagetnet)):
            execute_query(connection, "INSERT INTO getnet (data, valor) VALUES ('{}', '{}');".format(datagetnet[i], valorgetnet[i]))

        pasta.close()
        os.system('taskkill /im Excel.exe')
        await asyncio.sleep(0)