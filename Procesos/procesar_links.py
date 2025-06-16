# procesar_links.py
from crear_carpetas import crear_estructura_carpetas, crear_carpetas_unidas
import pandas as pd
import re
import sys
import os
import subprocess
import time
import shutil
import PyPDF2  # Asegúrate de instalar esta biblioteca con `pip install PyPDF2`

def esperar_descarga_completa(nombre_archivo_base, timeout=30):
    carpeta_descargas = os.path.join(os.path.expanduser('~'), 'Downloads')
    archivo_crdownload = os.path.join(carpeta_descargas, f"{nombre_archivo_base}.crdownload")
    tiempo_espera = 0

    while os.path.exists(archivo_crdownload) and tiempo_espera < timeout:
        print("⏳ Esperando que termine la descarga...")
        time.sleep(1)
        tiempo_espera += 1

    return not os.path.exists(archivo_crdownload)

def es_pdf_valido(ruta_pdf):
    try:
        with open(ruta_pdf, 'rb') as f:
            reader = PyPDF2.PdfReader(f)
            return len(reader.pages) > 0  # Verifica que el PDF tenga al menos una página
    except Exception:
        return False

def mover_ultimo_pdf_a_carpeta(destino):
    carpeta_descargas = os.path.join(os.path.expanduser('~'), 'Downloads')
    archivos_pdf = [f for f in os.listdir(carpeta_descargas) if f.endswith('.pdf')]

    if not archivos_pdf:
        print("❌ No se encontraron archivos PDF en Descargas.")
        return None

    archivos_pdf.sort(key=lambda f: os.path.getctime(os.path.join(carpeta_descargas, f)), reverse=True)
    archivo_mas_reciente = archivos_pdf[0]
    nombre_base = os.path.splitext(archivo_mas_reciente)[0]

    if not esperar_descarga_completa(nombre_base):
        print("❌ Descarga no se completó dentro del tiempo esperado.")
        return None

    ruta_origen = os.path.join(carpeta_descargas, archivo_mas_reciente)
    ruta_destino = os.path.join(destino, archivo_mas_reciente)

    # Verificar que el PDF sea válido
    if not es_pdf_valido(ruta_origen):
        print("❌ El archivo PDF está corrupto y no se moverá.")
        return None

    try:
        shutil.move(ruta_origen, ruta_destino)
        print(f"✅ PDF movido a: {ruta_destino}")
        return ruta_destino
    except Exception as e:
        print(f"❌ Error al mover archivo: {e}")
        return None

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

def mover_pdf_a_carpeta(ruta_pdf, destino):
    # Verificar que el PDF sea válido
    if not es_pdf_valido(ruta_pdf):
        print("❌ El archivo PDF está corrupto y no se moverá.")
        return None

    try:
        shutil.move(ruta_pdf, destino)
        print(f"✅ PDF movido a: {destino}")
        return destino
    except Exception as e:
        print(f"❌ Error al mover archivo: {e}")
        return None

def buscar_y_descargar_links(archivo_excel, columna, navegador, fila_desde, fila_hasta, columna_carpeta, columna_carpeta_extra):
    ruta_destino = crear_estructura_carpetas()

    if archivo_excel.endswith('.xlsx'):
        df = pd.read_excel(archivo_excel)
    elif archivo_excel.endswith('.csv'):
        df = pd.read_csv(archivo_excel)
    else:
        print("❌ Formato no soportado.")
        return

    if columna not in df.columns or columna_carpeta not in df.columns or columna_carpeta_extra not in df.columns:
        print(f"❌ Las columnas '{columna}', '{columna_carpeta}' o '{columna_carpeta_extra}' no existen.")
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

    lista_parte1 = df.loc[fila_desde:fila_hasta-1, columna_carpeta].apply(lambda x: str(int(x)) if isinstance(x, float) and x.is_integer() else str(x)).tolist()
    lista_parte2 = df.loc[fila_desde:fila_hasta-1, columna_carpeta_extra].astype(str).tolist()

    crear_carpetas_unidas(ruta_destino, lista_parte1, lista_parte2)


    link_pattern = re.compile(r'https://drive\.google\.com/[^\s,"]+')
    total_descargas = 0

    # Abre todos los enlaces simultáneamente
    for i, row in df_rango.iterrows():
        val = row[columna]
        links = link_pattern.findall(str(val))

        for link in links:
            nuevo_link = transformar_link(link)
            if nuevo_link:
                print(f"⬇️ Abriendo: {nuevo_link}")
                subprocess.Popen([exe_navegador, nuevo_link])

    print("⏳ Esperando 1 minuto para que se descarguen los PDFs...")
    time.sleep(60)  # Espera 60 segundos

    # Mueve los PDFs a las carpetas correspondientes
    carpeta_descargas = os.path.join(os.path.expanduser('~'), 'Downloads')
    archivos_pdf = [f for f in os.listdir(carpeta_descargas) if f.endswith('.pdf')]

    for i, row in df_rango.iterrows():
        valor1 = row[columna_carpeta]
        parte1 = str(int(valor1)) if isinstance(valor1, float) and valor1.is_integer() else str(valor1).strip()
        parte2 = str(row[columna_carpeta_extra]).strip()
        nombre_carpeta = f"{parte1} {parte2}".strip()
        destino_final = os.path.join(ruta_destino, nombre_carpeta)

        for archivo_pdf in archivos_pdf:
            if parte1.lower() in archivo_pdf.lower():
                ruta_pdf = os.path.join(carpeta_descargas, archivo_pdf)
                mover_pdf_a_carpeta(ruta_pdf, destino_final)
                total_descargas += 1

    if total_descargas == 0:
        print("⚠️ No se movió ningún archivo PDF.")
    else:
        print(f"✅ Se movieron {total_descargas} archivos PDF correctamente.")

     # 🔄 Mover carpetas a subcarpetas según segundo nombre
    print("\n🚚 Reubicando carpetas según el segundo nombre...")

    subcarpetas = ['Técnico', 'Tecnólogo', 'Auxiliar', 'Operario']
    carpetas_creadas = os.listdir(ruta_destino)

    for carpeta in carpetas_creadas:
        ruta_actual = os.path.join(ruta_destino, carpeta)
        if os.path.isdir(ruta_actual):
            partes = carpeta.strip().split()
            if len(partes) >= 2:
                segundo_nombre = partes[-1]
                if segundo_nombre in subcarpetas:
                    nueva_ruta = os.path.join(ruta_destino, segundo_nombre, carpeta)
                    try:
                        shutil.move(ruta_actual, nueva_ruta)
                        print(f"✅ {carpeta} movida a {segundo_nombre}")
                    except Exception as e:
                        print(f"❌ Error al mover {carpeta}: {e}")

if __name__ == "__main__":
    if len(sys.argv) < 8:
        print("Uso: python procesar_links.py <archivo_excel> <nombre_columna> <navegador> <fila_desde> <fila_hasta> <columna_carpeta> <columna_carpeta_extra>")
        sys.exit(1)

    archivo_excel = sys.argv[1]
    nombre_columna = sys.argv[2]
    navegador = sys.argv[3].lower()
    fila_desde = int(sys.argv[4]) - 2
    fila_hasta = int(sys.argv[5]) - 1
    columna_carpeta = sys.argv[6]
    columna_carpeta_extra = sys.argv[7]

    buscar_y_descargar_links(archivo_excel, nombre_columna, navegador, fila_desde, fila_hasta, columna_carpeta, columna_carpeta_extra)

