# main.py
# Menú principal para elegir entre Asistencia y Checklist

import tkinter as tk
from PIL import Image, ImageTk

def ejecutar_asistencia(root):
    root.destroy()
    __import__('main1').main()

def ejecutar_checklist(root):
    root.destroy()
    __import__('ejecutable').main()

def main():
    root = tk.Tk()
    root.title("Menú Principal")
    root.configure(bg="white")
    width, height = 450, 450
    screen_width = root.winfo_screenwidth()
    screen_height = root.winfo_screenheight()
    x = (screen_width - width) // 2
    y = (screen_height - height) // 2
    root.geometry(f"{width}x{height}+{x}+{y}")

    # Logo
    img_path = r"C:\\Users\\user\\Desktop\\OCAGLOBAL\\Firma\\Logo.png"
    img = Image.open(img_path)
    img = img.resize((180, 110))
    logo = ImageTk.PhotoImage(img)
    img_label = tk.Label(root, image=logo, bg="white")
    img_label.image = logo
    img_label.pack(pady=(20, 10))

    # Título estilizado
    label = tk.Label(root, text="Seleccione una opción:", font=("Segoe UI", 14, "bold"), bg="white", fg="#003366")
    label.pack(pady=10)

    # Botones estilizados
    btn_asistencia = tk.Button(root, text="Generar Reporte de Asistencia", width=30, font=("Segoe UI", 10), bg="#e0f0ff", activebackground="#c0e7ff", relief="flat", command=lambda: ejecutar_asistencia(root))
    btn_asistencia.pack(pady=10)

    btn_checklist = tk.Button(root, text="Generar Datos de Checklist", width=30, font=("Segoe UI", 10), bg="#e0f0ff", activebackground="#c0e7ff", relief="flat", command=lambda: ejecutar_checklist(root))
    btn_checklist.pack(pady=5)

    root.mainloop()

if __name__ == "__main__":
    main()
