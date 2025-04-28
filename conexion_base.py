# conexion.py
# Establece conexión con SQL Server usando pyodbc

import pyodbc
from config import server, database, username, password

def obtener_conexion():
    """
    Devuelve una conexión activa a la base de datos SQL Server,
    o None si ocurre algún error.
    """
    connection_string = (
        f'DRIVER={{ODBC Driver 17 for SQL Server}};'
        f'SERVER={server};DATABASE={database};UID={username};PWD={password}'
    )
    try:
        connection = pyodbc.connect(connection_string)
        return connection
    except pyodbc.Error as ex:
        print(f"Error de conexión: {ex.args[1]}")
        return None
