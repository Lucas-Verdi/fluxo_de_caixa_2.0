from cobranca_safra import *

def ler_cobranca_santander():
    global arquivo_cobranca_santander
    arquivo_cobranca_santander = filedialog.askopenfilename()
    label = Label(janela, text="{} CARREGADO".format(arquivo_cobranca_santander), font="Arial 7")
    label.grid(column=0, row=10)
    
async def cobranca_santander():
    global arquivo_cobranca_santander
    pasta = Book(arquivo_cobranca_santander)
    planilha = pasta.sheets[0]

    execute_query(connection, 'USE fluxodecaixa;')

    soma = 0
    planilha['1:6'].delete()
    lr = planilha.range('A2').end('down').row

    for i in range(2, lr + 1):
        data = planilha.range(f'D{i}').value
        data1 = planilha.range(f'D{i + 1}').value
        valor = planilha.range(f'C{i}').value
        convertido = valor.replace('.', '')
        convertido2 = convertido.replace(',', '.')
        soma += float(convertido2)
        if data != data1 or data1 == None:
            date = datetime.strptime(data, '%d/%m/%Y').date()
            data_cobranca_santander.append(date)
            valor_cobranca_santander.append(soma)
            soma = 0

    for i in range(0, len(data_cobranca_santander)):
        execute_query(connection, "INSERT INTO santander (data, valor) VALUES ('{}', '{}');".format(data_cobranca_santander[i], valor_cobranca_santander[i]))

    pasta.close()
    await asyncio.sleep(0)
