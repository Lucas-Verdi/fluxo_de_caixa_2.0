import datetime

import pyautogui
from getnet import *


def lersafrapay():
    global arquivosafrapay
    arquivosafrapay = filedialog.askopenfilename()
    label = Label(janela, text="{} CARREGADO".format(arquivosafrapay), font="Arial 7")
    label.grid(column=0, row=4)


async def safrapay():
    global arquivosafrapay
    if arquivosafrapay == None:
        print('Vazio')
    else:
        pasta = Book(arquivosafrapay)
        planilha = pasta.sheets[0]

        pasta.app.api.WindowState = -4137

        execute_query(connection, 'USE fluxodecaixa;')

        lr = planilha.range('F2').end('down').row

        planilha.range('F2:F{}'.format(lr)).api.Replace('.', '/')

        celula = planilha.range('A2').value
        if celula == None:
            print('VAZIO')
        else:
            pyautogui.moveTo(400, 0)
            pyautogui.click()
            planilha.range('F1:F{}'.format(lr)).select()
            pyautogui.sleep(0.5)
            pyautogui.hotkey('alt', 's')
            pyautogui.sleep(0.5)
            pyautogui.press('a')
            pyautogui.sleep(0.5)
            pyautogui.press('l')
            pyautogui.sleep(0.5)
            pyautogui.press('enter')
        
            soma = 0
            for i in range(2, lr + 1):
                data = planilha.range(f'F{i}').value
                data1 = planilha.range(f'F{i + 1}').value
                valor = planilha.range(f'L{i}').value
                soma += valor
                if data != data1:
                    date = planilha.range(f'F{i}').value
                    datasafrapay.append(date)
                    valorsafrapay.append(soma)
                    soma = 0

        for i in range(0, len(datasafrapay)):
            execute_query(connection, "INSERT INTO safra (data, valor) VALUES ('{}', '{}');".format(datasafrapay[i], valorsafrapay[i]))

        pasta.close()
        os.system('taskkill /im Excel.exe')
        await asyncio.sleep(0)