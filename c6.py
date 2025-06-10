from depositos import *

def ler_c6():
    global arquivo_c6
    arquivo_c6 = filedialog.askopenfilename()
    label = Label(janela, text="{} CARREGADO".format(arquivo_c6), font="Arial 7")
    label.grid(column=0, row=14)
    
async def c6():
    global arquivo_c6
    if arquivo_c6 == None:
        print('Vazio')
    else:
        pasta = Book(arquivo_c6)
        planilha = pasta.sheets[0]

        execute_query(connection, 'USE fluxodecaixa;')

        lr = planilha.range('A2').end('down').row

        soma = 0
        celula = planilha.range('A2').value
        if celula == None:
            print('VAZIO')
        else:
            for i in range(2, lr + 1):
                data = planilha.range(f'J{i}').value
                data1 = planilha.range(f'J{i + 1}').value
                valor = planilha.range(f'E{i}').value
                if isinstance(data, datetime):
                    soma += valor
                    if data != data1 or data1 == None:
                        data_c6.append(data)
                        valor_c6.append(soma)
                        print(soma)
                        soma = 0

        for i in range(0, len(data_c6)):
            execute_query(connection, "INSERT INTO c6 (data, valor) VALUES ('{}', '{}');".format(data_c6[i], valor_c6[i]))

        pasta.close()
        os.system('taskkill /im Excel.exe')
        await asyncio.sleep(0)