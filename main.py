from truncate import *
from threading import Thread

class Th(Thread):
    def __init__(self, num):
        Thread.__init__(self)
        self.num = num

    async def run_async(self):
        await getnet()
        await safrapay()
        await cobrancabb()
        await cobranca_safra()
        await cobranca_santander()
        await depositos()
        await despesas()
        await result()
        await truncate()

    def run(self):
        asyncio.run(self.run_async())
        


def start():
    a = Th(1)
    a.start()

janela.title('FLUXO DE CAIXA')
janela.geometry("500x550")

Label1 = Label(janela, text='Insira as pastas de trabalho:', font="Arial 10 bold", justify=CENTER)
Label1.grid(column=0, row=0, padx=150, pady=10)

Botao1 = Button(janela, text='GETNET', font="Arial 10")
Botao1.grid(column=0, row=1, padx=10, pady=10)
Botao1.bind("<Button>", lambda e: lergetnet())

Botao2 = Button(janela, text='SAFRAPAY', font="Arial 10")
Botao2.grid(column=0, row=3, padx=10, pady=10)
Botao2.bind("<Button>", lambda e: lersafrapay())

Botao3 = Button(janela, text='COBRANÇA BB', font="Arial 10")
Botao3.grid(column=0, row=5, padx=10, pady=10)
Botao3.bind("<Button>", lambda e: ler_cobrancabb())

Botao4 = Button(janela, text='COBRANÇA SAFRA', font="Arial 10")
Botao4.grid(column=0, row=7, padx=10, pady=10)
Botao4.bind("<Button>", lambda e: ler_cobranca_safra())

Botao5 = Button(janela, text='COBRANÇA SANTANDER', font="Arial 10")
Botao5.grid(column=0, row=9, padx=10, pady=10)
Botao5.bind("<Button>", lambda e: ler_cobranca_santander())

Botao6 = Button(janela, text='DEPOSITOS', font="Arial 10")
Botao6.grid(column=0, row=11, padx=10, pady=10)
Botao6.bind("<Button>", lambda e: ler_depositos())

Botao7 = Button(janela, text='DESPESAS', font="Arial 10")
Botao7.grid(column=0, row=13, padx=10, pady=10)
Botao7.bind("<Button>", lambda e: ler_despesas())

controle = Button(janela, text='GERAR CONTROLE', font="Arial 10")
controle.grid(column=0, row=15, padx=10, pady=10)
controle.bind("<Button>", lambda e: start())

janela.mainloop()
