import tkinter as tk
from tkinter import filedialog, messagebox, simpledialog
import pandas as pd
import sqlite3
import os
import re
from datetime import datetime

# Crear o conectar a la base de datos
def crear_bd():
    conn = sqlite3.connect('configuraciones.db')
    c = conn.cursor()
    c.execute('''CREATE TABLE IF NOT EXISTS settings 
                 (id INTEGER PRIMARY KEY, archivo TEXT, hoja TEXT, celda TEXT)''')
    conn.commit()
    conn.close()

# Guardar la configuración en la base de datos
def guardar_settings(settings):
    conn = sqlite3.connect('configuraciones.db')
    c = conn.cursor()
    c.executemany("INSERT INTO settings (archivo, hoja, celda) VALUES (?, ?, ?)", settings)
    conn.commit()
    conn.close()

# Cargar la configuración desde la base de datos
def cargar_settings():
    conn = sqlite3.connect('configuraciones.db')
    c = conn.cursor()
    c.execute("SELECT archivo, hoja, celda FROM settings")
    settings = c.fetchall()
    conn.close()
    return settings

# Seleccionar los archivos de Excel
def seleccionar_archivos():
    archivos = filedialog.askopenfilenames(
        title="Selecciona los archivos Excel",
        filetypes=[("Archivos de Excel", "*.xlsx *.xls")],
        multiple=True)
    if not archivos:
        messagebox.showwarning("Advertencia", "Debes seleccionar al menos un archivo")
        return None
    return archivos

# Convertir la referencia de celda (A1) a índices de fila y columna
def convertir_referencia(referencia):
    col_str = re.findall(r'[A-Z]+', referencia)[0]
    row_str = re.findall(r'\d+', referencia)[0]
    col = sum([(ord(char) - 64) * (26 ** i) for i, char in enumerate(reversed(col_str))]) - 1
    row = int(row_str) - 1
    return row, col

# Procesar el valor de una celda, rango o columna
def procesar_valor(sheet, referencia):
    if ":" in referencia:  # Rango de celdas
        return sheet.loc[:, referencia]
    elif re.match(r"[A-Z]+$", referencia):  # Columna entera
        return sheet[referencia]
    else:  # Celda única
        row, col = convertir_referencia(referencia)
        return sheet.iloc[row, col]

# Configurar las celdas y hojas a extraer
def configurar_celdas():
    configuracion = []
    while True:
        archivo = simpledialog.askstring("Entrada", "Ingrese el nombre del archivo (sin extensión) o 'fin' para terminar:")
        if archivo and archivo.lower() != "fin":
            hoja = simpledialog.askstring("Entrada", f"Ingrese el nombre de la hoja del archivo {archivo}:")
            if hoja:
                celdas = simpledialog.askstring("Entrada", f"Ingrese las celdas que desea extraer de la hoja {hoja} del archivo {archivo} (separadas por coma):")
                if celdas:
                    configuracion.append((archivo, hoja, celdas))
                else:
                    messagebox.showwarning("Advertencia", "Debes ingresar celdas válidas.")
                    return None
            else:
                messagebox.showwarning("Advertencia", "Debes ingresar un nombre de hoja válido.")
                return None
        elif archivo and archivo.lower() == "fin":
            break
        else:
            messagebox.showwarning("Advertencia", "Debes ingresar un nombre de archivo válido.")
            return None
    if configuracion:
        guardar_settings(configuracion)
    return configuracion

# Consolidar reportes
def consolidar_reportes(archivos, settings):
    if not archivos or not settings:
        messagebox.showerror("Error", "No se puede consolidar sin archivos o configuraciones válidas.")
        return

    data = []
    archivos_col = []
    hojas_col = []
    celdas_col = []
    fecha_actual = datetime.now().strftime("%d-%m-%Y")

    for archivo in archivos:
        archivo_nombre = os.path.basename(archivo).replace(".xlsx", "").replace(".xls", "")
        for setting in settings:
            if setting[0] in archivo_nombre:
                hoja = setting[1]
                df = pd.read_excel(archivo, sheet_name=hoja, header=None)  # Selecciona la hoja especificada
                celdas = setting[2].split(",")
                for celda in celdas:
                    valor = procesar_valor(df, celda.strip())
                    data.append(valor)
                    archivos_col.append(archivo_nombre)
                    hojas_col.append(hoja)
                    celdas_col.append(celda.strip())

    # Crear un DataFrame consolidado con las referencias y valores
    reporte_df = pd.DataFrame({
        "Archivo": archivos_col,
        "Hoja": hojas_col,
        "Celda": celdas_col,
        fecha_actual: data
    })

    # Guardar el reporte consolidado en un archivo
    reporte_df.to_excel("reporte_consolidado.xlsx", index=False)
    messagebox.showinfo("Éxito", "El reporte consolidado ha sido generado.")

# Estilo moderno para los botones
def estilo_moderno(widget):
    widget.config(
        bg="#4E4E4E",  # Fondo gris oscuro para contraste
        fg="#FFFFFF",  # Texto blanco
        font=("Helvetica", 12, "bold"),
        activebackground="#5E5E5E",  # Fondo al hacer clic
        activeforeground="#FFFFFF",  # Texto al hacer clic
        bd=0,  # Sin borde
        highlightthickness=0,  # Sin borde
        height=2,  # Altura consistente para todos los botones
        width=30,  # Ancho consistente para todos los botones
        pady=10,
        padx=10
    )

# Función principal
def main():
    # Verificar si existe la base de datos, si no, crearla
    if not os.path.exists('configuraciones.db'):
        crear_bd()

    # Configurar la ventana principal
    root = tk.Tk()
    root.title("Consolidación de Reportes")
    root.geometry("550x500")
    root.configure(bg="#2B2B2B")  # Fondo oscuro para la ventana principal

    # Título
    title = tk.Label(root, text="Consolidación de Reportes", font=("Helvetica", 16, "bold"), fg="#FFFFFF", bg="#2B2B2B")
    title.pack(pady=10)

    # Frame estilo "tarjeta" para las instrucciones
    instructions_frame = tk.Frame(root, bg="#3C3C3C", bd=1, relief="solid")
    instructions_frame.pack(pady=10, padx=10, fill="both", expand=False)

    # Instrucciones
    instrucciones = tk.Label(instructions_frame, text=(
        "Esta aplicación permite consolidar información de múltiples archivos Excel.\n"
        "Primero, selecciona los archivos de Excel. Luego puedes usar una configuración previa\n"
        "o definir nuevas celdas para extraer datos de cada archivo y hoja. Finalmente, consolida\n"
        "los reportes en un único archivo Excel. Puedes definir celdas individuales, rangos de\n"
        "celdas o columnas enteras."
    ), font=("Helvetica", 10), fg="#C0C0C0", bg="#3C3C3C", justify="left")
    instrucciones.pack(pady=10, padx=10)

    # Frame estilo "tarjeta" para los botones
    card_frame = tk.Frame(root, bg="#3C3C3C", bd=1, relief="solid")
    card_frame.pack(pady=10, padx=10, fill="both", expand=True)


    # Botón para ver configuraciones previas
    def ver_configuraciones_previas():
        settings = cargar_settings()
        if not settings:
            messagebox.showinfo("Configuraciones", "No se encontraron configuraciones previas.")
        else:
            detalles = "\n".join([f"Archivo: {archivo}, Hoja: {hoja}, Celdas: {celda}" for archivo, hoja, celda in settings])
            usar_previos = messagebox.askyesno("Configuraciones Previas", f"Se encontraron las siguientes configuraciones:\n\n{detalles}\n\n¿Desea usarlas?")
            if usar_previos:
                return settings
            else:
                return configurar_celdas()

    btn_ver_previos = tk.Button(card_frame, text="Ver Configuraciones Previas", command=lambda: ver_configuraciones_previas())
    estilo_moderno(btn_ver_previos)
    btn_ver_previos.pack(pady=5)

    # Botón para configurar nuevas celdas
    btn_configurar = tk.Button(card_frame, text="Configurar Nuevas Celdas", command=lambda: configurar_celdas())
    estilo_moderno(btn_configurar)
    btn_configurar.pack(pady=5)

    # Botón para consolidar reportes
    btn_consolidar = tk.Button(card_frame, text="Consolidar Reportes", command=lambda: consolidar_reportes(seleccionar_archivos(), cargar_settings()))
    estilo_moderno(btn_consolidar)
    btn_consolidar.pack(pady=5)

    # Copyright
    copyright_label = tk.Label(root, text="Franz Eckermann 2024", font=("Helvetica", 10), fg="#FFFFFF", bg="#2B2B2B")
    copyright_label.pack(side="bottom", pady=10)

    root.mainloop()

if __name__ == "__main__":
    main()
