import pandas as pd
from tkinter import messagebox, filedialog

class Funciones:
    def __init__(self, engine):
        self.engine = engine
        self.dataframe = pd.DataFrame()

    def consultar_checklist(self, checklist_id):
        try:
            conn = self.engine.raw_connection()
            cursor = conn.cursor()
            cursor.execute("EXEC dbo.GetCheckListData @CheckListId = ?", (checklist_id,))
            columns = [desc[0] for desc in cursor.description]
            rows = cursor.fetchall()
            cursor.close()
            conn.close()

            self.dataframe = pd.DataFrame([list(r) for r in rows], columns=columns)
            return self.dataframe
        except Exception as e:
            messagebox.showerror("Error", str(e))
            return pd.DataFrame()

    def exportar_excel(self, columnas, fecha_inicio=None, fecha_fin=None):
        try:
            if not columnas:
                messagebox.showwarning("Advertencia", "No se han seleccionado columnas.")
                return

            df_filtrado = self.dataframe[list(columnas)]

            if "FechaSistema" in df_filtrado.columns:
                df_filtrado["FechaSistema"] = pd.to_datetime(df_filtrado["FechaSistema"], errors='coerce')

                if fecha_inicio and not fecha_fin:
                    # ✅ FILTRO EXACTO para fecha única
                    fecha_unica = pd.to_datetime(fecha_inicio)
                    df_filtrado = df_filtrado[df_filtrado["FechaSistema"].dt.date == fecha_unica.date()]

                if fecha_inicio and fecha_fin:
                    fecha_inicio_dt = pd.to_datetime(fecha_inicio)
                    fecha_fin_dt = pd.to_datetime(fecha_fin)
                    df_filtrado = df_filtrado[
                        (df_filtrado["FechaSistema"] >= fecha_inicio_dt) &
                        (df_filtrado["FechaSistema"] <= fecha_fin_dt)
        ]
            ruta = filedialog.asksaveasfilename(
                defaultextension=".xlsx",
                filetypes=[("Excel files", "*.xlsx")]
            )
            if ruta:
                df_filtrado.to_excel(ruta, index=False)
                messagebox.showinfo("Éxito", f"Archivo exportado en:\n{ruta}")
        except Exception as e:
            messagebox.showerror("Error", str(e))
