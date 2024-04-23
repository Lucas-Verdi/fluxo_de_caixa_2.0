import xlwings

from despesas import *

async def result():
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

    cursor.execute("SELECT data, SUM(valor) FROM despesas GROUP BY data")
    despesas_tot = cursor.fetchall()


    data_atual = datetime.now()
    app = xlwings.App()
    pasta = app.books.add()
    planilha = pasta.sheets.active
    contdata = 3
    contdata2 = 1

    planilha.range('A1').value = "DATA"
    planilha.range('B1').value = "GETNET"
    planilha.range('C1').value = "SAFRAPAY"
    planilha.range('D1').value = "COBRANÇA BB"
    planilha.range('E1').value = "COBRANÇA SAFRA"
    planilha.range('F1').value = "COBRANÇA SANTANDER"
    planilha.range('G1').value = "DEPOSITOS"
    planilha.range('H1').value = "DESPESAS"

    planilha.range('A2').value = data_atual.strftime("%x")

    for i in range(0, 1460):
        incremento = data_atual + timedelta(days=contdata2)
        convertido = incremento.date()
        planilha.range('A{}'.format(contdata)).value = convertido.strftime("%x")
        contdata += 1
        contdata2 += 1

    await insertgetnet(planilha=planilha)
    await insertsafrapay(planilha=planilha)
    await insertcobrancabb(planilha=planilha)
    await insertcobrancasafra(planilha=planilha)
    await insertcobrancasantander(planilha=planilha)
    await insertdepositos(planilha=planilha)
    await insertdespesas(planilha=planilha)
    await asyncio.sleep(0)

async def insertgetnet(planilha):
    local = []
    lr = planilha.range('A2').end('down').row
    for i in range(0, len(getnet_tot)):
        data = getnet_tot[i][1]
        local.append(data)
    for i in range(2, lr + 1):
        data2 = planilha.range(f'A{i}').value
        date = data2.date()
        if date in local:
            indice = local.index(date)
            planilha.range(f'B{i}').value = getnet_tot[indice][2]
    await asyncio.sleep(0)

async def insertsafrapay(planilha):
    local = []
    lr = planilha.range('A2').end('down').row
    for i in range(0, len(safrapay_tot)):
        data = safrapay_tot[i][1]
        local.append(data)
    for i in range(2, lr + 1):
        data2 = planilha.range(f'A{i}').value
        date = data2.date()
        if date in local:
            indice = local.index(date)
            planilha.range(f'C{i}').value = safrapay_tot[indice][2]
    await asyncio.sleep(0)

async def insertcobrancabb(planilha):
    local = []
    lr = planilha.range('A2').end('down').row
    for i in range(0, len(cobranca_bb_tot)):
        data = cobranca_bb_tot[i][1]
        local.append(data)
    for i in range(2, lr + 1):
        data2 = planilha.range(f'A{i}').value
        date = data2.date()
        if date in local:
            indice = local.index(date)
            planilha.range(f'D{i}').value = cobranca_bb_tot[indice][2]
    await asyncio.sleep(0)

async def insertcobrancasafra(planilha):
    local = []
    lr = planilha.range('A2').end('down').row
    for i in range(0, len(cobranca_safra_tot)):
        data = cobranca_safra_tot[i][1]
        local.append(data)
    for i in range(2, lr + 1):
        data2 = planilha.range(f'A{i}').value
        date = data2.date()
        if date in local:
            indice = local.index(date)
            planilha.range(f'E{i}').value = cobranca_safra_tot[indice][2]
    await asyncio.sleep(0)

async def insertcobrancasantander(planilha):
    local = []
    lr = planilha.range('A2').end('down').row
    for i in range(0, len(cobranca_santander_tot)):
        data = cobranca_santander_tot[i][1]
        local.append(data)
    for i in range(2, lr + 1):
        data2 = planilha.range(f'A{i}').value
        date = data2.date()
        if date in local:
            indice = local.index(date)
            planilha.range(f'F{i}').value = cobranca_santander_tot[indice][2]
    await asyncio.sleep(0)

async def insertdepositos(planilha):
    local = []
    lr = planilha.range('A2').end('down').row
    for i in range(0, len(depositos_tot)):
        data = depositos_tot[i][1]
        local.append(data)
    for i in range(2, lr + 1):
        data2 = planilha.range(f'A{i}').value
        date = data2.date()
        if date in local:
            indice = local.index(date)
            planilha.range(f'G{i}').value = depositos_tot[indice][2]
    await asyncio.sleep(0)

async def insertdespesas(planilha):
    local = []
    lr = planilha.range('A2').end('down').row
    for i in range(0, len(despesas_tot)):
        data = despesas_tot[i][0]
        local.append(data)
    for i in range(2, lr + 1):
        data2 = planilha.range(f'A{i}').value
        date = data2.date()
        if date in local:
            indice = local.index(date)
            planilha.range(f'H{i}').value = despesas_tot[indice][1]
    await asyncio.sleep(0)


