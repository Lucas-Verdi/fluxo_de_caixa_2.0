from result import *

async def truncate():
    cursor.execute("TRUNCATE getnet;")
    cursor.execute("TRUNCATE safra;")
    cursor.execute("TRUNCATE bbrasil;")
    cursor.execute("TRUNCATE cobsafra;")
    cursor.execute("TRUNCATE santander;")
    cursor.execute("TRUNCATE depositos;")
    cursor.execute("TRUNCATE despesas;")
    await asyncio.sleep(0)