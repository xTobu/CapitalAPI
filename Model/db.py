import mysql.connector

def conn():
    conn = mysql.connector.connect(
    host = "127.0.0.1",
    user = "root",
    password = "0000",
    database = "future",
    )
    return conn