import pyodbc

def connect():
    global connection
    connection = pyodbc.connect(r'DRIVER={SQL SERVER};Server=DESKTOP-O8B5SCV\MSSQLSERVER1;Database=testpython;Trusted_Connection=yes;')


connect()