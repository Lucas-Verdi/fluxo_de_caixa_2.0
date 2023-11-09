from despesas import *

def query():
    global getnet_tot
    global safrapay_tot
    global cobranca_bb_tot
    global cobranca_safra_tot
    global cobranca_santander_tot
    global depositos_tot
    global despesas_tot

    cursor.execute("SELECT * FROM getnet ORDER BY data")
    getnet_tot = cursor.fetchall()

    cursor.execute("SELECT * FROM safra ORDER BY data")
    safrapay_tot = cursor.fetchall()

    cursor.execute("SELECT * FROM bbrasil ORDER BY data")
    cobranca_bb_tot = cursor.fetchall()

    cursor.execute("SELECT * FROM cobsafra ORDER BY data")
    cobranca_safra_tot = cursor.fetchall()

    cursor.execute("SELECT * FROM santander ORDER BY data")
    cobranca_santander_tot = cursor.fetchall()

    cursor.execute("SELECT * FROM depositos ORDER BY data")
    depositos_tot = cursor.fetchall()

    cursor.execute("SELECT * FROM despesas ORDER BY data")
    despesas_tot = cursor.fetchall()
