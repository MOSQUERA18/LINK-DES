import tkinter as tk
from tkinter import messagebox, filedialog
import subprocess
import sys
import os

def seleccionar_archivo():
    ruta = filedialog.askopenfilename(
        title="Seleccionar archivo Excel o CSV",
        filetypes=[("Archivos Excel", "*.xlsx"), ("Archivos CSV", "*.csv")]
    )
    if ruta:
        entry_archivo.delete(0, tk.END)
        entry_archivo.insert(0, ruta)

def navegador_seleccionado(opcion):
    # Mostrar sugerencia justo cuando el usuario seleccione un navegador
    messagebox.showinfo("Sugerencia", f"Asegúrate de que el navegador '{opcion}' esté abierto antes de continuar.")

def iniciar():
    archivo = entry_archivo.get().strip()
    columna = entry_columna.get().strip()
    navegador = navegador_var.get()
    fila_desde = entry_fila_desde.get().strip()
    fila_hasta = entry_fila_hasta.get().strip()
    # columna_carpeta = entry_columna_carpeta.get().strip()
    # columna_carpeta_extra = entry_columna_carpeta_extra.get().strip()

    if not archivo or not os.path.isfile(archivo):
        messagebox.showerror("Error", "Selecciona un archivo válido.")
        return
    if not columna:
        messagebox.showerror("Error", "Escribe el nombre exacto de la columna donde están los links.")
        return
    if navegador not in ["Firefox", "Chrome", "Edge"]:
        messagebox.showerror("Error", "Selecciona un navegador válido.")
        return
    if not fila_desde.isdigit() or not fila_hasta.isdigit():
        messagebox.showerror("Error", "Las filas deben ser números enteros.")
        return
    # if not columna_carpeta:
    #     messagebox.showerror("Error", "Escribe el nombre exacto de la columna para crear las carpetas.")
    #     return
    # if not columna_carpeta_extra:
    #     messagebox.showerror("Error", "Escribe el nombre de la columna de continuación para las carpetas.")
    #     return

    ruta_procesar = os.path.join(os.path.dirname(__file__), "procesar_links.py")
    subprocess.Popen([
        sys.executable, ruta_procesar, archivo, columna, navegador.lower(), fila_desde, fila_hasta
    ])


# Crear ventana
ventana = tk.Tk()
ventana.title("Abrir Links desde Excel")
ventana.geometry("600x400")
ventana.resizable(False, False)

# Estilo general
padding_args = {'padx': 10, 'pady': 5, 'sticky': 'w'}

# === ARCHIVO ===
frame_archivo = tk.LabelFrame(ventana, text="Archivo", padx=10, pady=10)
frame_archivo.pack(fill="x", padx=15, pady=10)

tk.Label(frame_archivo, text="Selecciona el archivo Excel o CSV:").grid(row=0, column=0, columnspan=2, **padding_args)
entry_archivo = tk.Entry(frame_archivo, width=50)
entry_archivo.grid(row=1, column=0, **padding_args)
btn_examinar = tk.Button(frame_archivo, text="Examinar", command=seleccionar_archivo)
btn_examinar.grid(row=1, column=1, **padding_args)

# === OPCIONES ===
frame_opciones = tk.LabelFrame(ventana, text="Opciones", padx=10, pady=10)
frame_opciones.pack(fill="x", padx=15, pady=10)

tk.Label(frame_opciones, text="Columna con los links:").grid(row=0, column=0, **padding_args)
entry_columna = tk.Entry(frame_opciones)
entry_columna.grid(row=0, column=1, **padding_args)

# tk.Label(frame_opciones, text="Columna para carpetas:").grid(row=1, column=0, **padding_args)
# entry_columna_carpeta = tk.Entry(frame_opciones)
# entry_columna_carpeta.grid(row=1, column=1, **padding_args)

# tk.Label(frame_opciones, text="Columna continuación carpeta:").grid(row=1, column=2, **padding_args)
# entry_columna_carpeta_extra = tk.Entry(frame_opciones)
# entry_columna_carpeta_extra.grid(row=1, column=3, **padding_args)

tk.Label(frame_opciones, text="Navegador:").grid(row=2, column=0, **padding_args)
navegador_var = tk.StringVar()
navegador_var.set("Chrome")

# Crear menú de opciones con función callback
navegador_opciones = tk.OptionMenu(
    frame_opciones,
    navegador_var,
    "Firefox",
    "Chrome",
    "Edge",
    command=navegador_seleccionado  # <-- Esta es la clave
)
navegador_opciones.grid(row=2, column=1, **padding_args)

tk.Label(frame_opciones, text="Fila desde (ej. 2):").grid(row=3, column=0, **padding_args)
entry_fila_desde = tk.Entry(frame_opciones, width=10)
entry_fila_desde.grid(row=3, column=1, **padding_args)

tk.Label(frame_opciones, text="Fila hasta (ej. 20):").grid(row=4, column=0, **padding_args)
entry_fila_hasta = tk.Entry(frame_opciones, width=10)
entry_fila_hasta.grid(row=4, column=1, **padding_args)

# === BOTÓN INICIAR ===
frame_boton = tk.Frame(ventana)
frame_boton.pack(pady=15)
tk.Button(frame_boton, text="Iniciar", command=iniciar, width=20, bg="#4CAF50", fg="white").pack()

ventana.mainloop()
