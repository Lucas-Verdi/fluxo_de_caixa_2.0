from depositos import *

def ler_despesas():
    global arquivo_despesas
    arquivo_despesas = filedialog.askopenfilename()
    label = Label(janela, text="{} CARREGADO".format(arquivo_despesas), font="Arial 7")
    label.grid(column=0, row=14)
    
def despesas():
    global arquivo_despesas
    pasta = Book(arquivo_despesas)
    planilha = pasta.sheets[0]

    execute_query(connection, 'USE fluxodecaixa;')

    planilha['1:5'].delete()
    lr = planilha.range('A2').end('down').row

    soma = 0
    for i in range(2, lr + 1):
        data = planilha.range(f'A{i}').value
        data1 = planilha.range(f'A{i + 1}').value
        valor = planilha.range(f'G{i}').value
        soma += valor
        if data != data1 or data1 == None:
            data_despesas.append(data)
            valor_despesas.append(soma)
            soma = 0

    for i in range(0, len(data_despesas)):
        execute_query(connection, "INSERT INTO despesas (data, valor) VALUES ('{}', '{}');".format(data_despesas[i], valor_despesas[i]))

    pasta.close()