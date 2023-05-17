import pyodbc

def connect():
    global connection
    connection = pyodbc.connect(r'DRIVER={SQL SERVER};Server=(yourserver);Database=(yourdatabase);Trusted_Connection=yes;')


connect()
