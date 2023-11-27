from truncate import *
from threading import Thread

pb = ttk.Progressbar(
    janela,
    orient='horizontal',
    mode='determinate',
    length=300
)

class Th(Thread):
    def __init__(self, num):
        Thread.__init__(self)
        self.num = num
        Th.daemon = True

    async def run_async(self):
        pb['value'] = 0
        await getnet()
        pb['value'] = 12
        await safrapay()
        pb['value'] = 24
        await cobrancabb()
        pb['value'] = 36
        await cobranca_safra()
        pb['value'] = 48
        await cobranca_santander()
        pb['value'] = 60
        await depositos()
        pb['value'] = 72
        await despesas()
        pb['value'] = 84
        await result()
        pb['value'] = 100
        await truncate()

    def run(self):
        asyncio.run(self.run_async())
        


def start():
    a = Th(1)
    a.start()

janela.title('FLUXO DE CAIXA')
janela.geometry("500x600")

Label1 = Label(janela, text='Insira as pastas de trabalho:', font="Arial 10 bold", justify=CENTER)
Label1.grid(column=0, row=0, padx=150, pady=10)

Botao1 = ttk.Button(janela, text='GETNET')
Botao1.grid(column=0, row=1, padx=10, pady=10)
Botao1.bind("<Button>", lambda e: lergetnet())

Botao2 = ttk.Button(janela, text='SAFRAPAY')
Botao2.grid(column=0, row=3, padx=10, pady=10)
Botao2.bind("<Button>", lambda e: lersafrapay())

Botao3 = ttk.Button(janela, text='COBRANÇA BB')
Botao3.grid(column=0, row=5, padx=10, pady=10)
Botao3.bind("<Button>", lambda e: ler_cobrancabb())

Botao4 = ttk.Button(janela, text='COBRANÇA SAFRA')
Botao4.grid(column=0, row=7, padx=10, pady=10)
Botao4.bind("<Button>", lambda e: ler_cobranca_safra())

Botao5 = ttk.Button(janela, text='COBRANÇA SANTANDER')
Botao5.grid(column=0, row=9, padx=10, pady=10)
Botao5.bind("<Button>", lambda e: ler_cobranca_santander())

Botao6 = ttk.Button(janela, text='DEPOSITOS')
Botao6.grid(column=0, row=11, padx=10, pady=10)
Botao6.bind("<Button>", lambda e: ler_depositos())

Botao7 = ttk.Button(janela, text='DESPESAS')
Botao7.grid(column=0, row=13, padx=10, pady=10)
Botao7.bind("<Button>", lambda e: ler_despesas())

controle = ttk.Button(janela, text='GERAR CONTROLE')
controle.grid(column=0, row=15, padx=10, pady=10)
controle.bind("<Button>", lambda e: start())

pb.grid(column=0, row=16, padx=10, pady=10)

janela.mainloop()
