# conexion/conexion_sql.py
import urllib
from sqlalchemy import create_engine

def get_engine():
    try:
        server = '40.75.98.215'
        database = 'SafeSmart_Checklist'
        username = 'jarbildo_dev'
        password = 'oca$2022'
        driver = 'ODBC Driver 17 for SQL Server'

        params = urllib.parse.quote_plus(
            f"DRIVER={{{driver}}};SERVER={server};DATABASE={database};UID={username};PWD={password}"
        )
        engine = create_engine(f"mssql+pyodbc:///?odbc_connect={params}")
        return engine
    except Exception as e:
        print(f"❌ Error de conexión: {e}")
        return None
