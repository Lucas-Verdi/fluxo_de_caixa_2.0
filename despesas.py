from depositos import *

def ler_despesas():
    global arquivo_despesas
    arquivo_despesas = filedialog.askopenfilename(multiple=True)
    label = Label(janela, text="{} CARREGADO".format(arquivo_despesas[0]), font="Arial 7")
    label.grid(column=0, row=14)
    
async def despesas():
    await arquivo1()
    await arquivo2()
    await arquivo3()
    await arquivo4()

async def arquivo1():
    global arquivo_despesas
    pasta = Book(arquivo_despesas[0])
    planilha = pasta.sheets[0]

    execute_query(connection, 'USE fluxodecaixa;')

    planilha['1:5'].delete()
    lr = planilha.range('A2').end('down').row

    soma = 0
    for i in range(2, lr + 1):
        data = planilha.range(f'A{i}').value
        data1 = planilha.range(f'A{i + 1}').value
        valor = planilha.range(f'G{i}').value
        if isinstance(data, datetime):
            soma += valor
            if data != data1 or data1 == None:
                data_despesas.append(data)
                valor_despesas.append(soma)
                soma = 0

    for i in range(0, len(data_despesas)):
        execute_query(connection, "INSERT INTO despesas (data, valor) VALUES ('{}', '{}');".format(data_despesas[i], valor_despesas[i]))

    pasta.close()
    os.system('taskkill /im Excel.exe')
    await asyncio.sleep(0)

async def arquivo2():
    global arquivo_despesas
    pasta = Book(arquivo_despesas[1])
    planilha = pasta.sheets[0]

    execute_query(connection, 'USE fluxodecaixa;')

    planilha['1:5'].delete()
    lr = planilha.range('A2').end('down').row

    soma = 0
    for i in range(2, lr + 1):
        data = planilha.range(f'A{i}').value
        data1 = planilha.range(f'A{i + 1}').value
        valor = planilha.range(f'G{i}').value
        if isinstance(data, datetime):
            soma += valor
            if data != data1 or data1 == None:
                data_despesas2.append(data)
                valor_despesas2.append(soma)
                soma = 0

    for i in range(0, len(data_despesas2)):
        execute_query(connection, "INSERT INTO despesas (data, valor) VALUES ('{}', '{}');".format(data_despesas2[i], valor_despesas2[i]))

    pasta.close()
    os.system('taskkill /im Excel.exe')
    await asyncio.sleep(0)

async def arquivo3():
    global arquivo_despesas
    pasta = Book(arquivo_despesas[2])
    planilha = pasta.sheets[0]

    execute_query(connection, 'USE fluxodecaixa;')

    planilha['1:5'].delete()
    lr = planilha.range('A2').end('down').row

    soma = 0
    for i in range(2, lr + 1):
        data = planilha.range(f'A{i}').value
        data1 = planilha.range(f'A{i + 1}').value
        valor = planilha.range(f'G{i}').value
        if isinstance(data, datetime):
            soma += valor
            if data != data1 or data1 == None:
                data_despesas3.append(data)
                valor_despesas3.append(soma)
                soma = 0

    for i in range(0, len(data_despesas3)):
        execute_query(connection, "INSERT INTO despesas (data, valor) VALUES ('{}', '{}');".format(data_despesas3[i], valor_despesas3[i]))

    pasta.close()
    os.system('taskkill /im Excel.exe')
    await asyncio.sleep(0)

async def arquivo4():
    global arquivo_despesas
    pasta = Book(arquivo_despesas[3])
    planilha = pasta.sheets[0]

    execute_query(connection, 'USE fluxodecaixa;')

    planilha['1:5'].delete()
    lr = planilha.range('A2').end('down').row

    soma = 0
    for i in range(2, lr + 1):
        data = planilha.range(f'A{i}').value
        data1 = planilha.range(f'A{i + 1}').value
        valor = planilha.range(f'G{i}').value
        if isinstance(data, datetime):
            soma += valor
            if data != data1 or data1 == None:
                data_despesas4.append(data)
                valor_despesas4.append(soma)
                soma = 0

    for i in range(0, len(data_despesas4)):
        execute_query(connection, "INSERT INTO despesas (data, valor) VALUES ('{}', '{}');".format(data_despesas4[i], valor_despesas4[i]))

    pasta.close()
    os.system('taskkill /im Excel.exe')
    await asyncio.sleep(0)