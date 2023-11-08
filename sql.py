import mysql.connector
from mysql.connector import Error


connection = mysql.connector.connect(
    host='localhost',
    user='root',
    passwd='wolf'
)

def execute_query(connection, query):
    cursor = connection.cursor()
    try:
        cursor.execute(query)
        connection.commit()
        print("Query successful")
    except Error as err:
        print(f"Error: '{err}'")
