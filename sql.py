import mysql.connector
from mysql.connector import Error


connection = mysql.connector.connect(
    host='localhost',
    user='root',
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
