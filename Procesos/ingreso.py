import webbrowser
import time
import sys

# Verifica si se pasó la URL y el navegador como argumentos
if len(sys.argv) < 3:
    print("Uso: python ingreso.py <url> <navegador>")
    sys.exit(1)

url = sys.argv[1]
navegador = sys.argv[2].lower()

# Rutas comunes a navegadores en Windows
rutas = {
    "firefox": "C:\\Program Files\\Mozilla Firefox\\firefox.exe",
    "chrome": "C:\\Program Files\\Google\\Chrome\\Application\\chrome.exe",
    "edge": "C:\\Program Files (x86)\\Microsoft\\Edge\\Application\\msedge.exe"
}

# Validar navegador
if navegador not in rutas:
    print(f"El navegador '{navegador}' no está soportado.")
    sys.exit(1)

# Registrar el navegador con su ruta
webbrowser.register(navegador, None, webbrowser.BackgroundBrowser(rutas[navegador]))

# Obtener el navegador registrado y abrir la URL
navegador_registrado = webbrowser.get(navegador)
navegador_registrado.open(url)

# Esperar para que el usuario vea la página
time.sleep(10)
