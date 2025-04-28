# ejecutable.py
import tkinter as tk
from conexion.conexion_sql import get_engine
from logica.funciones import Funciones
from vista.interfaz import Interfaz

def main():
    engine = get_engine()
    if engine:
        funciones = Funciones(engine)
        root = tk.Tk()
        root.configure(bg="white")
        app = Interfaz(root, funciones)
        root.mainloop()
    else:
        print("No se pudo establecer conexión a la base de datos.")

if __name__ == "__main__":
    main()


