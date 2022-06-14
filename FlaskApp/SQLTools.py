import time
import mysql.connector
from mysql.connector import Error

def create_server_connection(host_name, user_name, user_password, db_name):
    connection = None
    try:
        connection = mysql.connector.connect(
            host=host_name,
            user=user_name,
            passwd=user_password,
            database=db_name
        )
        print("MySQL Database connection successful")
    except Error as err:
        print(f"Error: '{err}'")

    return connection

def execute_query(connection, query):
    '''Save start time'''
    start_time = time.time()
    cursor = connection.cursor(buffered=True)
    try:
        cursor.execute(query)
        connection.commit()
        print("Query successful. Execution time: {} seconds".format(round(time.time() - start_time, 5)))
        
    except Error as err:
        print(f"Error: '{err}'")

    result = cursor.fetchall()
    return result