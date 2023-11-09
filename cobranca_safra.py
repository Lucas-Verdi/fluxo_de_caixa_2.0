from cobranca_bb import *

def ler_cobranca_safra():
    global arquivo_cobranca_safra
    arquivo_cobranca_safra = filedialog.askopenfilename()
    label = Label(janela, text="{} CARREGADO".format(arquivo_cobranca_safra), font="Arial 7")
    label.grid(column=0, row=8)
    
def cobranca_safra():
    global arquivo_cobranca_safra
    pasta = Book(arquivo_cobranca_safra)
    planilha = pasta.sheets[0]

    execute_query(connection, 'USE fluxodecaixa;')

    soma = 0
    planilha['1:5'].delete()
    lr = planilha.range('A2').end('down').row
    
    for i in range(2, lr + 1):
        data = planilha.range(f'A{i}').value
        data1 = planilha.range(f'A{i + 1}').value
        valor = planilha.range(f'I{i}').value
        convertido = valor.replace('.', '')
        convertido2 = convertido.replace(',', '.')
        soma += float(convertido2)
        if data != data1 or data1 == None:
            date = datetime.strptime(data, '%d/%m/%Y').date()
            data_cobranca_safra.append(date)
            valor_cobranca_safra.append(soma)
            soma = 0

    for i in range(0, len(data_cobranca_safra)):
        execute_query(connection, "INSERT INTO cobsafra (data, valor) VALUES ('{}', '{}');".format(data_cobranca_safra[i], valor_cobranca_safra[i]))

    pasta.close()