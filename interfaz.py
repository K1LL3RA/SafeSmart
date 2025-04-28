# interfaz.py
# Interfaz gráfica del programa hecha con Tkinter
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from config import plantilla_path
from conexion_base import obtener_conexion
from consulta_sql import ejecutar_consulta
from excel_utils import llenar_excel
import os
import pandas as pd
from datetime import datetime
from calendar import month_name
import psutil

MES_NOMBRES_ES = {
    1: "Enero", 2: "Febrero", 3: "Marzo", 4: "Abril", 5: "Mayo", 6: "Junio",
    7: "Julio", 8: "Agosto", 9: "Septiembre", 10: "Octubre", 11: "Noviembre", 12: "Diciembre"
}
def cerrar_excel_si_abierto(nombre_archivo):
    for proc in psutil.process_iter(['pid', 'name', 'open_files']):
        try:
            if proc.info['name'] and 'excel' in proc.info['name'].lower():
                for file in proc.info['open_files'] or []:
                    if nombre_archivo.lower() in file.path.lower():
                        proc.kill()
                        return True
        except (psutil.NoSuchProcess, psutil.AccessDenied):
            continue
    return False

def main():
    root = tk.Tk()
    root.title("Generar Registro en Excel")
    root.configure(bg="white")
    
    style = ttk.Style()
    style.theme_use('clam')
    style.configure("TLabel", background="white", font=("Segoe UI", 10))
    style.configure("TCombobox", padding=4)
    style.configure("TButton", font=("Segoe UI", 10), padding=6, relief="flat", background="#e0f0ff")
    style.map("TButton", background=[("active", "#c0e7ff")])

    window_width = 1000
    window_height = 600
    screen_width = root.winfo_screenwidth()
    screen_height = root.winfo_screenheight()
    x_cordinate = int((screen_width/2) - (window_width/2))
    y_cordinate = int((screen_height/2) - (window_height/2))
    root.geometry(f"{window_width}x{window_height}+{x_cordinate}+{y_cordinate}")

    checklist_id = ttk.Entry(root, width=40)
    anio_seleccionado = tk.StringVar()
    mes_seleccionado = tk.StringVar()
    dia_seleccionado = tk.StringVar()
    fecha_seleccionada = tk.StringVar()
    texto_g50_entry = ttk.Entry(root, width=40)
    texto_g52_entry = ttk.Entry(root, width=40)
    ruta_imagen = tk.StringVar()
    carpeta_destino = tk.StringVar()
    carpeta_firmas = tk.StringVar()

    temas_tratados_listbox = tk.Listbox(root, selectmode='multiple', exportselection=0, width=50, height=8, font=("Segoe UI", 9))
    horas_inicio_listbox = tk.Listbox(root, selectmode='multiple', exportselection=0, width=30, height=5, font=("Segoe UI", 9))

    anio_combo = ttk.Combobox(root, textvariable=anio_seleccionado, state='disabled')
    mes_combo = ttk.Combobox(root, textvariable=mes_seleccionado, state='disabled')
    dia_combo = ttk.Combobox(root, textvariable=dia_seleccionado, state='disabled')

    def seleccionar_todos_los_temas():
        temas_tratados_listbox.select_set(0, tk.END)

    def seleccionar_todas_las_horas():
        horas_inicio_listbox.select_set(0, tk.END)

    def seleccionar_carpeta_firmas():
        ruta = filedialog.askdirectory(title="Selecciona la carpeta de firmas")
        if ruta:
            carpeta_firmas.set(ruta)
            messagebox.showinfo("Carpeta de firmas seleccionada", ruta)

    def seleccionar_imagen():
        ruta = filedialog.askopenfilename(title="Seleccionar Imagen", filetypes=[("Imágenes", "*.png;*.jpg;*.jpeg")])
        ruta_imagen.set(ruta)
        if ruta:
            messagebox.showinfo("Imagen seleccionada", ruta)

    def seleccionar_carpeta():
        carpeta = filedialog.askdirectory()
        if carpeta:
            carpeta_destino.set(carpeta)

    def actualizar_horas_y_temas(*args):
        data = getattr(root, 'data', None)
        if data is not None and anio_seleccionado.get() and mes_seleccionado.get() and dia_seleccionado.get():
            mes_num = root.mes_nombre_a_numero.get(mes_seleccionado.get(), 0)
            fecha_final = f"{int(dia_seleccionado.get()):02}/{int(mes_num):02}/{anio_seleccionado.get()}"
            fecha_seleccionada.set(fecha_final)
            filtrado = data[data['Fecha'] == fecha_final]
            horas = filtrado['Hora Inicio'].dropna().unique().tolist()
            temas = filtrado['Tema Tratado'].dropna().unique().tolist()
            horas_inicio_listbox.delete(0, tk.END)
            temas_tratados_listbox.delete(0, tk.END)
            for hora in horas:
                horas_inicio_listbox.insert(tk.END, hora)
            for tema in temas:
                temas_tratados_listbox.insert(tk.END, tema)

    def cargar_datos():
        if not checklist_id.get().isdigit():
            messagebox.showerror("Error", "Ingrese un CheckListId válido.")
            return

        conn = obtener_conexion()
        if not conn:
            messagebox.showerror("Error", "No se pudo conectar a la base de datos.")
            return

        data = ejecutar_consulta(conn, checklist_id.get())
        conn.close()

        if data is None or data.empty:
            messagebox.showerror("Error", "No se encontraron datos para ese ID.")
            return

        root.data = data
        fechas = pd.to_datetime(data['Fecha'].unique(), dayfirst=True)
        anios = sorted(set(f.year for f in fechas))
        anio_combo.config(values=anios, state='readonly')

        def actualizar_meses(*args):
            año = int(anio_seleccionado.get())
            meses_nums = sorted(set(f.month for f in fechas if f.year == año))
            meses_nombres = [MES_NOMBRES_ES[m] for m in meses_nums]
            mes_combo.config(values=meses_nombres, state='readonly')
            mes_combo.grid()
            root.mes_nombre_a_numero = dict(zip(meses_nombres, meses_nums))
            dia_combo.grid_remove()
            root.update_idletasks()
            root.geometry("")

        def actualizar_dias(*args):
            año = int(anio_seleccionado.get())
            mes_nombre = mes_seleccionado.get()
            mes = root.mes_nombre_a_numero.get(mes_nombre, 0)
            dias = sorted(set(f.day for f in fechas if f.year == año and f.month == mes))
            dia_combo.config(values=dias, state='readonly')
            dia_combo.grid()
            root.update_idletasks()
            root.geometry("")

        anio_seleccionado.trace_add('write', lambda *args: actualizar_meses())
        mes_seleccionado.trace_add('write', lambda *args: actualizar_dias())
        dia_seleccionado.trace_add('write', actualizar_horas_y_temas)
        messagebox.showinfo("Éxito", "Datos cargados correctamente. Seleccione año, mes y día.")

    
    def generar_excel():
        data = getattr(root, 'data', None)
        if data is None:
            messagebox.showerror("Error", "Debe cargar los datos primero.")
            return

        fecha = fecha_seleccionada.get()
        horas = [horas_inicio_listbox.get(i) for i in horas_inicio_listbox.curselection()]
        temas = [temas_tratados_listbox.get(i) for i in temas_tratados_listbox.curselection()]

        if not fecha or not horas or not temas:
            messagebox.showerror("Error", "Seleccione fecha, hora y tema.")
            return

        try:
            data['Fecha'] = pd.to_datetime(data['Fecha'], dayfirst=True).dt.strftime("%d/%m/%Y")
        except:
            messagebox.showerror("Error", "No se pudo convertir el campo 'Fecha'.")
            return

        filtrado = data[(data['Fecha'] == fecha) &
                        (data['Hora Inicio'].isin(horas)) &
                        (data['Tema Tratado'].isin(temas))]
        root.filtrado = filtrado

        if filtrado.empty:
            messagebox.showerror("Error", "No hay datos para la selección realizada.")
            return

        if not carpeta_destino.get():
            messagebox.showerror("Error", "Seleccione una carpeta de destino.")
            return

        try:
            clasificacion_registro = filtrado['Clasificación del Registro'].unique()
            clasificacion_nombre = "_".join(clasificacion_registro).replace(' ', '_')
        except:
            clasificacion_nombre = "Checklist"

        try:
            fecha_formateada = datetime.strptime(fecha, "%d/%m/%Y").strftime("%d-%m-%Y")
        except:
            fecha_formateada = fecha.replace('/', '-')

        nombre_archivo = f"{clasificacion_nombre}_{fecha_formateada}.xlsx"
        ruta_archivo = os.path.join(carpeta_destino.get(), nombre_archivo)

        llenar_excel(filtrado, ruta_archivo, plantilla_path,
                 texto_g50_entry.get(), texto_g52_entry.get(), ruta_imagen.get(), carpeta_firmas.get(), redimension_firma=True)
        messagebox.showinfo("Éxito", f"Archivo generado: {ruta_archivo}")
        btn_pdf.config(state="normal")

    def generar_pdf():
        try:
            import win32com.client
            import time

            if not carpeta_destino.get():
                messagebox.showerror("Error", "No se ha seleccionado carpeta de destino.")
                return

            fecha = fecha_seleccionada.get()
            clasificacion_registro = root.filtrado['Clasificación del Registro'].unique()
            clasificacion_nombre = "_".join(clasificacion_registro).replace(' ', '_')
            fecha_formateada = datetime.strptime(fecha, "%d/%m/%Y").strftime("%d-%m-%Y")
            nombre_archivo = f"{clasificacion_nombre}_{fecha_formateada}.xlsx"
            ruta_excel = os.path.join(carpeta_destino.get(), nombre_archivo)
            ruta_pdf = os.path.splitext(ruta_excel)[0] + ".pdf"

            if not os.path.exists(ruta_excel):
                messagebox.showerror("Error", f"El archivo Excel no existe:\n{ruta_excel}")
                return

            for _ in range(3):
                try:
                    with open(ruta_excel, 'a'):
                        break
                except PermissionError:
                    time.sleep(1)
            else:
                messagebox.showerror("Error", "El archivo está en uso. Cierra Excel o reinicia el programa.")
                return

            excel = win32com.client.DispatchEx("Excel.Application")
            excel.Visible = False
            wb = None

            try:
                wb = excel.Workbooks.Open(ruta_excel, ReadOnly=1)
                wb.ExportAsFixedFormat(0, ruta_pdf)
                messagebox.showinfo("Éxito", f"PDF generado correctamente:\n{ruta_pdf}")
            except Exception as e:
                messagebox.showerror("Error", f"No se pudo generar el PDF:\n{e}")
            finally:
                if wb:
                    wb.Close(False)
                excel.Quit()

        except Exception as e:
            messagebox.showerror("Error", f"Error inesperado:\n{e}")

    # Frame para botones de exportación
    frame_generar = tk.Frame(root, bg="white")
    frame_generar.grid(row=9, column=2, padx=10, pady=10, sticky="e")

    tt_btn_excel = ttk.Button(frame_generar, text="📄 Generar Excel", command=generar_excel, width=18)
    tt_btn_excel.pack(side="left", padx=5)

    btn_pdf = ttk.Button(frame_generar, text="📄 Generar PDF", command=generar_pdf, width=18, state="disabled")
    btn_pdf.pack(side="left", padx=5)


    # UI Placement
    ttk.Label(root, text="Ingresar CheckListId:").grid(row=0, column=0, padx=10, pady=5, sticky="e")
    checklist_id.grid(row=0, column=1, padx=10, pady=5, sticky="w")
    ttk.Button(root, text="🔍 Cargar Datos", command=cargar_datos).grid(row=0, column=2, padx=10, pady=5)

    ttk.Label(root, text="Seleccionar Año:").grid(row=1, column=0, padx=10, pady=5, sticky="e")
    anio_combo.grid(row=1, column=1, padx=10, pady=5, sticky="w")

    ttk.Label(root, text="Seleccionar Mes:").grid(row=2, column=0, padx=10, pady=5, sticky="e")
    mes_combo.grid(row=2, column=1, padx=10, pady=5, sticky="w")
    mes_combo.grid_remove()

    ttk.Label(root, text="Seleccionar Día:").grid(row=3, column=0, padx=10, pady=5, sticky="e")
    dia_combo.grid(row=3, column=1, padx=10, pady=5, sticky="w")
    dia_combo.grid_remove()

    ttk.Label(root, text="Seleccionar Temas:").grid(row=4, column=0, padx=10, pady=5, sticky="e")
    temas_tratados_listbox.grid(row=4, column=1, padx=10, pady=5, sticky="w")
    ttk.Button(root, text="📌 Seleccionar Todo", command=seleccionar_todos_los_temas).grid(row=4, column=2, padx=10, pady=5)

    ttk.Label(root, text="Seleccionar Horas:").grid(row=5, column=0, padx=10, pady=5, sticky="e")
    horas_inicio_listbox.grid(row=5, column=1, padx=10, pady=5, sticky="w")
    ttk.Button(root, text="⏰ Seleccionar Todo", command=seleccionar_todas_las_horas).grid(row=5, column=2, padx=10, pady=5)

    ttk.Label(root, text="Responsable del Registro:").grid(row=6, column=0, padx=10, pady=5, sticky="e")
    texto_g50_entry.grid(row=6, column=1, padx=10, pady=5, sticky="w")

    ttk.Label(root, text="Cargo:").grid(row=7, column=0, padx=10, pady=5, sticky="e")
    texto_g52_entry.grid(row=7, column=1, padx=10, pady=5, sticky="w")

    frame_seleccion = tk.Frame(root, bg="white")
    frame_seleccion.grid(row=8, column=0, columnspan=2, padx=10, pady=10, sticky="w")
    ttk.Button(frame_seleccion, text="🗂️ Seleccionar Carpeta", command=seleccionar_carpeta).pack(side="left", padx=5)
    ttk.Button(frame_seleccion, text="🖼️ Seleccionar Imagen", command=seleccionar_imagen).pack(side="left", padx=5)
    ttk.Button(frame_seleccion, text="✍️ Seleccionar carpeta de firmas", command=seleccionar_carpeta_firmas).pack(side="left", padx=5)

    root.mainloop()