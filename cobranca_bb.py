from safrapay import *

def ler_cobrancabb():
    global arquivo_cobrancabb
    arquivo_cobrancabb = filedialog.askopenfilename()
    label = Label(janela, text="{} CARREGADO".format(arquivo_cobrancabb), font="Arial 7")
    label.grid(column=0, row=6)


async def cobrancabb():
    global arquivo_cobrancabb
    if arquivo_cobrancabb == None:
        print('Vazio')
    else:
        pasta = Book(arquivo_cobrancabb)
        planilha = pasta.sheets[0]

        execute_query(connection, 'USE fluxodecaixa;')

        lr = planilha.range('A2').end('down').row

        celula = planilha.range('A2').value
        if celula == None:
            print('VAZIO')
        else:
            soma = 0
            for i in range(3, lr + 1):
                data = planilha.range(f'A{i}').value
                data1 = planilha.range(f'A{i + 1}').value
                valor = planilha.range(f'E{i}').value
                if isinstance(data, datetime):
                    soma += valor
                    if data != data1 or data1 == None:
                        data_cobrancabb.append(data)
                        valor_cobrancabb.append(soma)
                        soma = 0

        for i in range(0, len(data_cobrancabb)):
            execute_query(connection, "INSERT INTO bbrasil (data, valor) VALUES ('{}', '{}');".format(data_cobrancabb[i], valor_cobrancabb[i]))

        pasta.close()
        os.system('taskkill /im Excel.exe')
        await asyncio.sleep(0)