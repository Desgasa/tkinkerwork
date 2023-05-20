import mysql.connector
from mysql.connector import Error

def connect():
    global dbconn 
    dbconn = mysql.connector.connect(
    host = "yourhost",
    port = yourport, 
    user = "youruser",
    password = "",
    database = "yourdatabase"
)

connect()
