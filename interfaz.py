import tkinter as tk
from tkinter import messagebox, Listbox, SINGLE, ttk
from logica.funciones import Funciones
from datetime import datetime
import pandas as pd

class Interfaz:

    def __init__(self, root, funciones: Funciones):
        self.funciones = funciones
        self.root = root
        self.root.title("📊 Exportador de Checklist")
        self.root.geometry("1000x650")
        self.usar_rango = tk.BooleanVar()
        self.root.configure(bg="white")

        self.estilizar_widgets()
        self.crear_componentes()
        self.centrar_ventana()

    def estilizar_widgets(self):
        style = ttk.Style()
        style.theme_use('clam')
        style.configure("TButton", font=("Segoe UI", 10), padding=6, relief="flat", background="#e0f0ff")
        style.map("TButton",
                  background=[("active", "#c0e7ff")])
        style.configure("TLabel", background="white", font=("Segoe UI", 10))
        style.configure("TCombobox", padding=4)

    def crear_componentes(self):
        top_frame = tk.Frame(self.root, bg="white")
        top_frame.pack(pady=10)

        ttk.Label(top_frame, text="Ingrese el CheckList ID:").grid(row=0, column=0, padx=5)
        self.entry_id = ttk.Entry(top_frame, width=15)
        self.entry_id.grid(row=0, column=1, padx=5)
        ttk.Button(top_frame, text="🔍 Consultar", command=self.consultar).grid(row=0, column=2, padx=5)

        fecha_frame = tk.Frame(self.root, bg="white")
        fecha_frame.pack(pady=5)

        ttk.Label(fecha_frame, text="Año:").grid(row=0, column=0, padx=5)
        self.combo_anio = ttk.Combobox(fecha_frame, state="readonly", width=8)
        self.combo_anio.grid(row=0, column=1, padx=5)
        self.combo_anio.bind("<<ComboboxSelected>>", self.actualizar_meses)

        ttk.Label(fecha_frame, text="Mes:").grid(row=0, column=2, padx=5)
        self.combo_mes = ttk.Combobox(fecha_frame, state="readonly", width=5)
        self.combo_mes.grid(row=0, column=3, padx=5)
        self.combo_mes.bind("<<ComboboxSelected>>", self.actualizar_dias)

        ttk.Label(fecha_frame, text="Día:").grid(row=0, column=4, padx=5)
        self.combo_dia = ttk.Combobox(fecha_frame, state="readonly", width=5)
        self.combo_dia.grid(row=0, column=5, padx=5)

        ttk.Checkbutton(
            fecha_frame, text="Fecha única",
            variable=self.usar_rango, onvalue=True, offvalue=False,
            command=self.toggle_fecha_fin
        ).grid(row=0, column=6, padx=10)

        ttk.Label(fecha_frame, text="Año fin:").grid(row=1, column=0, padx=5)
        self.combo_anio_fin = ttk.Combobox(fecha_frame, state="readonly", width=8)
        self.combo_anio_fin.grid(row=1, column=1, padx=5)
        self.combo_anio_fin.bind("<<ComboboxSelected>>", self.actualizar_meses_fin)

        ttk.Label(fecha_frame, text="Mes fin:").grid(row=1, column=2, padx=5)
        self.combo_mes_fin = ttk.Combobox(fecha_frame, state="readonly", width=5)
        self.combo_mes_fin.grid(row=1, column=3, padx=5)
        self.combo_mes_fin.bind("<<ComboboxSelected>>", self.actualizar_dias_fin)

        ttk.Label(fecha_frame, text="Día fin:").grid(row=1, column=4, padx=5)
        self.combo_dia_fin = ttk.Combobox(fecha_frame, state="readonly", width=5)
        self.combo_dia_fin.grid(row=1, column=5, padx=5)

        self.toggle_fecha_fin()

        ttk.Label(self.root, text="Selecciona columnas y ordénalas:").pack(pady=(10, 0))

        frame = tk.Frame(self.root, bg="white")
        frame.pack(pady=10)

        self.listbox_disponibles = Listbox(frame, selectmode=SINGLE, width=40, height=20, font=("Segoe UI", 9))
        self.listbox_disponibles.grid(row=0, column=0, padx=10)

        btn_frame = tk.Frame(frame, bg="white")
        btn_frame.grid(row=0, column=1)
        ttk.Button(btn_frame, text="➕ Agregar ➜", command=self.agregar_columna).pack(pady=5)
        ttk.Button(btn_frame, text="📋 Agregar Todos ➜", command=self.agregar_todos).pack(pady=5)
        ttk.Button(btn_frame, text="⬅ Quitar", command=self.quitar_columna).pack(pady=5)

        self.listbox_seleccionadas = Listbox(frame, selectmode=SINGLE, width=40, height=20, font=("Segoe UI", 9))
        self.listbox_seleccionadas.grid(row=0, column=2, padx=10)

        orden_frame = tk.Frame(self.root, bg="white")
        orden_frame.pack(pady=5)
        ttk.Button(orden_frame, text="🔼 Subir", command=self.mover_arriba).grid(row=0, column=0, padx=10)
        ttk.Button(orden_frame, text="🔽 Bajar", command=self.mover_abajo).grid(row=0, column=1, padx=10)

        ttk.Button(self.root, text="📁 Exportar a Excel", command=self.exportar).pack(pady=10)
        ttk.Button(self.root, text="🔙 Volver al Menú Principal", command=lambda: [self.root.destroy(), __import__('main').main()]).pack(pady=10)

   
    def toggle_fecha_fin(self):
        estado = "disabled" if self.usar_rango.get() else "readonly"
        self.combo_anio_fin.configure(state=estado)
        self.combo_mes_fin.configure(state=estado)
        self.combo_dia_fin.configure(state=estado)

    def consultar(self):
        try:
            checklist_id = int(self.entry_id.get())
            df = self.funciones.consultar_checklist(checklist_id)

            if df.empty:
                messagebox.showinfo("Sin resultados", "No se encontraron datos.")
                return

            self.listbox_disponibles.delete(0, tk.END)
            self.listbox_seleccionadas.delete(0, tk.END)
            for col in df.columns:
                self.listbox_disponibles.insert(tk.END, col)

            if 'FechaSistema' in df.columns:
                fechas = pd.to_datetime(df['FechaSistema'], errors='coerce').dropna()
                self.fechas_filtradas = fechas

                anios = sorted(fechas.dt.year.unique())
                for combo in [self.combo_anio, self.combo_anio_fin]:
                    combo['values'] = anios
                    combo.set('')

                for combo in [self.combo_mes, self.combo_mes_fin, self.combo_dia, self.combo_dia_fin]:
                    combo.set('')
                    combo['values'] = []
        except ValueError:
            messagebox.showerror("Error", "Por favor ingrese un ID válido.")

    def actualizar_meses(self, event):
        anio = self.combo_anio.get()
        if anio:
            meses = self.fechas_filtradas[self.fechas_filtradas.dt.year == int(anio)].dt.month.unique()
            self.combo_mes['values'] = sorted(meses)

    def actualizar_dias(self, event):
        anio = self.combo_anio.get()
        mes = self.combo_mes.get()
        if anio and mes:
            dias = self.fechas_filtradas[
                (self.fechas_filtradas.dt.year == int(anio)) &
                (self.fechas_filtradas.dt.month == int(mes))
            ].dt.day.unique()
            self.combo_dia['values'] = sorted(dias)

    def actualizar_meses_fin(self, event):
        anio = self.combo_anio_fin.get()
        if anio:
            meses = self.fechas_filtradas[self.fechas_filtradas.dt.year == int(anio)].dt.month.unique()
            self.combo_mes_fin['values'] = sorted(meses)

    def actualizar_dias_fin(self, event):
        anio = self.combo_anio_fin.get()
        mes = self.combo_mes_fin.get()
        if anio and mes:
            dias = self.fechas_filtradas[
                (self.fechas_filtradas.dt.year == int(anio)) &
                (self.fechas_filtradas.dt.month == int(mes))
            ].dt.day.unique()
            self.combo_dia_fin['values'] = sorted(dias)

    def exportar(self):
        columnas = self.listbox_seleccionadas.get(0, tk.END)
        anio = self.combo_anio.get()
        mes = self.combo_mes.get()
        dia = self.combo_dia.get()
        fecha_inicio = None
        if anio and mes and dia:
            fecha_inicio = f"{anio}-{int(mes):02d}-{int(dia):02d}"

        fecha_fin = None
        if not self.usar_rango.get():  # ✅ Solo usar fecha_fin si NO es fecha única
            anio_f = self.combo_anio_fin.get()
            mes_f = self.combo_mes_fin.get()
            dia_f = self.combo_dia_fin.get()
            if anio_f and mes_f and dia_f:
                fecha_fin = f"{anio_f}-{int(mes_f):02d}-{int(dia_f):02d}"

        self.funciones.exportar_excel(columnas, fecha_inicio, fecha_fin)

    def agregar_columna(self):
        selected = self.listbox_disponibles.curselection()
        if selected:
            col = self.listbox_disponibles.get(selected[0])
            if col not in self.listbox_seleccionadas.get(0, tk.END):
                self.listbox_seleccionadas.insert(tk.END, col)

    def agregar_todos(self):
        disponibles = self.listbox_disponibles.get(0, tk.END)
        ya_seleccionadas = self.listbox_seleccionadas.get(0, tk.END)
        for col in disponibles:
            if col not in ya_seleccionadas:
                self.listbox_seleccionadas.insert(tk.END, col)

    def quitar_columna(self):
        selected = self.listbox_seleccionadas.curselection()
        if selected:
            self.listbox_seleccionadas.delete(selected[0])

    def mover_arriba(self):
        idx = self.listbox_seleccionadas.curselection()
        if not idx or idx[0] == 0:
            return
        texto = self.listbox_seleccionadas.get(idx)
        self.listbox_seleccionadas.delete(idx)
        self.listbox_seleccionadas.insert(idx[0] - 1, texto)
        self.listbox_seleccionadas.selection_set(idx[0] - 1)

    def mover_abajo(self):
        idx = self.listbox_seleccionadas.curselection()
        if not idx or idx[0] == self.listbox_seleccionadas.size() - 1:
            return
        texto = self.listbox_seleccionadas.get(idx)
        self.listbox_seleccionadas.delete(idx)
        self.listbox_seleccionadas.insert(idx[0] + 1, texto)
        self.listbox_seleccionadas.selection_set(idx[0] + 1)

    def centrar_ventana(self):
        self.root.update_idletasks()
        ancho = self.root.winfo_width()
        alto = self.root.winfo_height()
        pantalla_ancho = self.root.winfo_screenwidth()
        pantalla_alto = self.root.winfo_screenheight()
        x = (pantalla_ancho // 2) - (ancho // 2)
        y = (pantalla_alto // 2) - (alto // 2)
        self.root.geometry(f"{ancho}x{alto}+{x}+{y}")
