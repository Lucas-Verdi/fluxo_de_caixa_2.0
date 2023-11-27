from cobranca_santander import *

def ler_depositos():
    global arquivo_depositos
    arquivo_depositos = filedialog.askopenfilename()
    label = Label(janela, text="{} CARREGADO".format(arquivo_depositos), font="Arial 7")
    label.grid(column=0, row=12)
    
async def depositos():
    global arquivo_depositos
    pasta = Book(arquivo_depositos)
    planilha = pasta.sheets[0]

    execute_query(connection, 'USE fluxodecaixa;')

    lr = planilha.range('A2').end('down').row

    soma = 0
    for i in range(2, lr + 1):
        data = planilha.range(f'H{i}').value
        data1 = planilha.range(f'H{i + 1}').value
        valor = planilha.range(f'I{i}').value
        soma += valor
        if data != data1 or data1 == None:
            data_depositos.append(data)
            valor_depositos.append(soma)
            soma = 0

    for i in range(0, len(data_depositos)):
        execute_query(connection, "INSERT INTO depositos (data, valor) VALUES ('{}', '{}');".format(data_depositos[i], valor_depositos[i]))

    pasta.close()
    os.system('taskkill /im Excel.exe')
    await asyncio.sleep(0)