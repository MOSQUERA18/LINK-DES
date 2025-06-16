import pandas as pd
import re
import sys
import os
import subprocess
import time

def transformar_link(link):
    match = re.search(r'/d/([a-zA-Z0-9_-]+)', link)
    if match:
        file_id = match.group(1)
        print(f"📎 ID extraído: {file_id}")
        return f'https://drive.google.com/uc?export=download&id={file_id}'
    else:
        match_alt = re.search(r'id=([a-zA-Z0-9_-]+)', link)
        if match_alt:
            file_id = match_alt.group(1)
            print(f"📎 ID alternativo extraído: {file_id}")
            return f'https://drive.google.com/uc?export=download&id={file_id}'
    print("❌ No se pudo transformar el link:", link)
    return None

def buscar_y_descargar_links(archivo_excel, columna, navegador, fila_desde, fila_hasta):
    if archivo_excel.endswith('.xlsx'):
        df = pd.read_excel(archivo_excel)
    elif archivo_excel.endswith('.csv'):
        df = pd.read_csv(archivo_excel)
    else:
        print("❌ Formato no soportado.")
        return

    if columna not in df.columns:
        print(f"❌ La columna '{columna}' no existe.")
        return

    rutas = {
        "chrome": "C:\\Program Files\\Google\\Chrome\\Application\\chrome.exe",
        "firefox": "C:\\Program Files\\Mozilla Firefox\\firefox.exe",
        "edge": "C:\\Program Files (x86)\\Microsoft\\Edge\\Application\\msedge.exe"
    }

    if navegador not in rutas:
        print(f"❌ Navegador no soportado.")
        return

    exe_navegador = rutas[navegador]
    df_rango = df.iloc[fila_desde:fila_hasta]

    link_pattern = re.compile(r'https://drive\.google\.com/[^\s,"]+')
    total_descargas = 0

    for i, row in df_rango.iterrows():
        val = row[columna]
        links = link_pattern.findall(str(val))

        for link in links:
            nuevo_link = transformar_link(link)
            if nuevo_link:
                print(f"⬇️ Abriendo: {nuevo_link}")
                subprocess.Popen([exe_navegador, nuevo_link])
                total_descargas += 1

    print("⏳ Esperando 1 minuto para que se descarguen los PDFs...")
    time.sleep(60)

    if total_descargas == 0:
        print("⚠️ No se detectó ningún enlace válido.")
    else:
        print(f"✅ Se abrieron {total_descargas} enlaces para descarga.")

if __name__ == "__main__":
    if len(sys.argv) < 6:
        print("Uso: python procesar_links.py <archivo_excel> <nombre_columna> <navegador> <fila_desde> <fila_hasta>")
        sys.exit(1)

    archivo_excel = sys.argv[1]
    nombre_columna = sys.argv[2]
    navegador = sys.argv[3].lower()
    fila_desde = int(sys.argv[4]) - 1
    fila_hasta = int(sys.argv[5])

    buscar_y_descargar_links(archivo_excel, nombre_columna, navegador, fila_desde, fila_hasta)
