import os

def crear_estructura_carpetas():
    carpeta_principal = os.path.join(os.path.expanduser('~'), 'Downloads', 'ARCHIVOS TITULADA')

    try:
        if not os.path.exists(carpeta_principal):
            os.makedirs(carpeta_principal)
            print(f"📁 Carpeta principal creada: {carpeta_principal}")
        else:
            print(f"📂 Carpeta principal ya existe: {carpeta_principal}")


        
        return carpeta_principal

    except Exception as e:
        print(f"❌ Error al crear las carpetas: {e}")
        return None

""" def crear_carpetas_por_fila(ruta_base, nombres):
    for nombre in nombres:
        ruta_carpeta = os.path.join(ruta_base, str(nombre).strip())
        try:
            if not os.path.exists(ruta_carpeta):
                os.makedirs(ruta_carpeta)
                print(f"📁 Carpeta creada: {ruta_carpeta}")
            else:
                print(f"🔄 Carpeta ya existe: {ruta_carpeta}")
        except Exception as e:
            print(f"❌ Error al crear carpeta '{nombre}': {e}")"""

def crear_carpetas_unidas(ruta_base, lista_parte1, lista_parte2):
    """
    Crea carpetas uniendo elemento a elemento de dos listas, separando con espacio.
    """
    for parte1, parte2 in zip(lista_parte1, lista_parte2):
        nombre_carpeta = f"{str(parte1).strip()} {str(parte2).strip()}".strip()
        ruta_carpeta = os.path.join(ruta_base, nombre_carpeta)
        try:
            if not os.path.exists(ruta_carpeta):
                os.makedirs(ruta_carpeta)
                print(f"📁 Carpeta creada: {ruta_carpeta}")
            else:
                print(f"🔄 Carpeta ya existe: {ruta_carpeta}")
        except Exception as e:
            print(f"❌ Error al crear carpeta '{nombre_carpeta}': {e}")

if __name__ == "__main__":
    base = crear_estructura_carpetas()
    if base:
        # Ejemplo simple
        """crear_carpetas_por_fila(base, ["Ejemplo1", "Ejemplo2", "Ejemplo3"])"""

        # Ejemplo con nombres unidos
        crear_carpetas_unidas(base, ["Juan", "Ana"], ["Perez", "Lopez"])