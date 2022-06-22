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

def execute_queryNoTime(connection, query):
    cursor = connection.cursor(buffered=True)
    try:
        cursor.execute(query)
        connection.commit()
    except Error as err:
        print(f"Error: '{err}'")
    result = cursor.fetchall()
    return result

host = 'localhost'
user_name = 'taran'
user_password = 'DESKTOP-RMV78RR'
db_name = 'seekapp'


def testSpeed():
    companyID = 92
    startDate = '2022-05-01'
    endDate = '2022-06-31'
    
    groupBy = "level"
    connection = create_server_connection(host, user_name, user_password, db_name)

    query1 = "select id,question,option1,option2,option3,option4,option5,employees_type_data,actual_schedule_time from surveyquestions where compid = {} and actual_schedule_time >= '{}' and actual_schedule_time <= '{}' and scheduled_count<>0;".format(companyID, startDate, endDate)
    result1 = execute_query(connection, query1)

    query3 = "select {},count(*) from employee where companyId={} and status='active' group by {};".format(groupBy, companyID, groupBy)
    result3 = execute_query(connection, query3)

    allIDs = [46207, 46208, 46209, 46210, 46211, 46212, 46213, 46214, 46215, 46216, 46217, 46218, 46219, 46220, 46221, 46222, 46223, 46224, 46225, 46226, 46227, 46228, 46229, 46230, 46231, 46232, 46289, 46290, 46291, 46292, 46371, 46372, 46373, 46374, 46811, 46812, 46813, 46814, 46857, 46858, 46859, 46860, 46861, 46862, 46863, 46864, 46867, 46868, 46869, 46870, 47096, 47158, 47159, 47160, 47162, 47167, 47322, 47337, 47415, 47416, 47417, 47418, 47419, 47420, 47421, 47422, 47423, 47424, 47425, 47426, 47427, 47428, 47429, 47430, 47524, 47531, 47539, 47553, 48521, 48522, 48523, 48524, 48525, 48526, 48527, 48528, 48529, 48530, 48531, 48532, 48718, 48719, 48720, 48721, 48722, 48968, 48969, 48970, 48971, 48989, 49023, 49037]
    for questionID in allIDs:
        query2 = "select e.{},n.answer from notifications n left join employee e on n.empId=e.id where message_id={} and answer is not null;".format(groupBy, questionID)
        result2 = execute_query(connection, query2)

testSpeed()
