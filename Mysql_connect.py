import mysql.connector
from mysql.connector import Error
from mysql.connector import errorcode

try:
    connection = mysql.connector.connect(host='localhost',
                                         database='bikroy',
                                         user='root',
                                         password='')

    mySql_insert_query = """INSERT INTO info (binfo) 
                           VALUES 
                           ('hello im using this for testing purpose') """

    cursor = connection.cursor()
    result = cursor.execute(mySql_insert_query)
    connection.commit()
    print("Record inserted successfully into Laptop table")
    cursor.close()

except mysql.connector.Error as error:
    print("Failed to insert record into Laptop table {}".format(error))

finally:
    if (connection.is_connected()):
        connection.close()
        print("MySQL connection is closed")


"""

loop insert query 


cur = db.cursor()

myList = [1,2,3,4]
for row in myList:
   print row
   cur.execute("INSERT INTO table (a) VALUES (%s)" % row)


"""