import mysql.connector
from mysql.connector import Error


connection = mysql.connector.connect(
    host='solident.ddns.net',
    port='3408',
    user='wolf',
    passwd='wolf',
)

cursor = connection.cursor()
def execute_query(connection, query):
    try:
        cursor.execute(query)
        connection.commit()
        print("Query successful")
    except Error as err:
        print(f"Error: '{err}'")
