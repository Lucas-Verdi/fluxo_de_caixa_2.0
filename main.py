from cobranca_bb import *

def main():
    getnet()
    safrapay()
    cobrancabb()

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

controle = Button(janela, text='GERAR CONTROLE', font="Arial 10")
controle.grid(column=0, row=7, padx=10, pady=10)
controle.bind("<Button>", lambda e: main())

janela.mainloop()
