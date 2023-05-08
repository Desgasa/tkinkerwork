import mysql.connector
from mysql.connector import Error

def connect():
    global dbconn 
    dbconn = mysql.connector.connect(
    host = "localhost",
    port = 3306, 
    user = "root",
    password = "",
    database = "test2python"
)

connect()