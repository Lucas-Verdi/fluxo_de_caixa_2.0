from truncate import *
from tkinter import messagebox


def truncate_manual():
    cursor.execute('USE fluxodecaixa;')
    cursor.execute("TRUNCATE getnet;")
    cursor.execute("TRUNCATE safra;")
    cursor.execute("TRUNCATE bbrasil;")
    cursor.execute("TRUNCATE cobsafra;")
    cursor.execute("TRUNCATE santander;")
    cursor.execute("TRUNCATE depositos;")
    cursor.execute("TRUNCATE despesas;")
    messagebox.showinfo("Alerta", "Sucesso!")