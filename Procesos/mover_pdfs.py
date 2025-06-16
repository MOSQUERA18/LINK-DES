# mover_pdfs.py
import os
import shutil
import sys
import pandas as pd

def mover_pdfs(archivo_excel, columna_carpeta, fila_desde, fila_hasta):
    carpeta_descargas = os.path.join(os.path.expanduser("~"), "Downloads")
    carpeta_destino = os.path.join(carpeta_descargas, "ARCHIVOS TITULADA")

    if not os.path.exists(archivo_excel):
        print(f"❌ El archivo '{archivo_excel}' no existe.")
        return

    if archivo_excel.endswith('.xlsx'):
        df = pd.read_excel(archivo_excel)
    elif archivo_excel.endswith('.csv'):
        df = pd.read_csv(archivo_excel)
    else:
        print("❌ Formato no soportado.")
        return

    if columna_carpeta not in df.columns:
        print(f"❌ La columna '{columna_carpeta}' no existe en el archivo.")
        return

    df_rango = df.iloc[fila_desde:fila_hasta]
    nombres = df_rango[columna_carpeta].dropna().unique()

    archivos = [f for f in os.listdir(carpeta_descargas) if f.lower().endswith(".pdf")]

    if not archivos:
        print("⚠️ No se encontraron archivos PDF en la carpeta de descargas.")
        return

    for nombre in nombres:
        subcarpeta = os.path.join(carpeta_destino, str(nombre))
        if not os.path.exists(subcarpeta):
            os.makedirs(subcarpeta)

        for archivo in archivos:
            origen = os.path.join(carpeta_descargas, archivo)
            destino = os.path.join(subcarpeta, archivo)

            # Solo mover si aún no está en la carpeta destino
            if not os.path.exists(destino):
                try:
                    shutil.move(origen, destino)
                    print(f"📁 Movido: {archivo} → {subcarpeta}")
                except Exception as e:
                    print(f"❌ Error al mover {archivo}: {e}")

if __name__ == "__main__":
    if len(sys.argv) < 5:
        print("Uso: python mover_pdfs.py <archivo_excel> <columna_carpeta> <fila_desde> <fila_hasta>")
        sys.exit(1)

    archivo_excel = sys.argv[1]
    columna_carpeta = sys.argv[2]
    fila_desde = int(sys.argv[3]) - 1
    fila_hasta = int(sys.argv[4])

    mover_pdfs(archivo_excel, columna_carpeta, fila_desde, fila_hasta)
