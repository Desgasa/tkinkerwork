import mysql.connector
from mysql.connector import Error

def connect():
    global dbconn 
    dbconn = mysql.connector.connect(
        host = "localhost",
        port = 3306, 
        user = "root",
        password = "",
        database = "testpython"
)

def database():
        global result
        mycursor = dbconn.cursor()
        mycursor.execute(""" SELECT `value_text` FROM `ClientCode` WHERE `Treatyid` = "1" """)
        result = mycursor.fetchone()


connect()