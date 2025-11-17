import os
import time
import subprocess
import ctypes
import logging
import pyperclip
import pyautogui
from watchdog.observers import Observer
from watchdog.events import FileSystemEventHandler
import pygetwindow as gw
import json


# ============================================================
#  RUTA FIJA DONDE EL INSTALADOR CREARÁ CONFIG.JSON
# ============================================================

CONFIG_DIR = r"C:\Program Files\visorirys"
# CONFIG_DIR = r"D:\mundomedicodental\code"
CONFIG_FILE = os.path.join(CONFIG_DIR, "config.json")


# ============================================================
#  CARGAR CONFIGURACIÓN
# ============================================================

def load_config():
    """
    Carga config.json desde la ruta fija que genera el instalador.
    """
    if not os.path.exists(CONFIG_FILE):
        raise FileNotFoundError(
            f"ERROR: No se encontró config.json en: {CONFIG_FILE}"
        )

    with open(CONFIG_FILE, "r", encoding="utf-8") as f:
        return json.load(f)


config = load_config()

# Datos del config.json
carpeta_export = config.get("carpeta_export", "")
folderirys = config.get("folderirys", "")

ruta_visor = os.path.join(folderirys, "NNT.exe")


# ============================================================
#  CONFIGURACIONES DEL PROGRAMA
# ============================================================

ESPERA_USB_SEC = 6
REINTENTOS = 2
REINTENTO_DELAY = 3
LOG_FILE = os.path.join(CONFIG_DIR, "watch_nnt.log")

VENTANAS_MINIMIZAR = ["Neowise"]


logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
    handlers=[
        logging.FileHandler(LOG_FILE, encoding="utf-8"),
        logging.StreamHandler()
    ]
)


# ============================================================
#  FUNCIONES DE UTILIDAD
# ============================================================

def esperar_por_usb(timeout=ESPERA_USB_SEC):
    """Espera al archivo NNT.ini indicando que USB/dongle está listo."""
    start = time.time()
    while time.time() - start < timeout:
        if os.path.exists(os.path.join(folderirys, "NNT.ini")):
            return True
        time.sleep(0.5)
    return False


def lanzar_nnt():
    """Inicia NNT.exe con reintentos."""
    if not os.path.isfile(ruta_visor):
        return False, "No existe NNT.exe en la ruta configurada"

    if not esperar_por_usb():
        logging.warning("Timeout esperando driver/dongle. Intentando igual...")

    try:
        proc = subprocess.Popen([ruta_visor], cwd=folderirys)
        logging.info(f"NNT.exe lanzado (pid={proc.pid}) – esperando carga inicial…")
        time.sleep(8)
        return True, f"popen (pid={proc.pid})"
    except Exception as e:
        logging.warning(f"Popen falló: {e}")

    return False, "Todos los métodos fallaron"


def maximizar_y_traer_irys():
    """Trae la ventana iRYS al frente y minimiza otras."""
    todas = gw.getAllWindows()
    ventana_irys = next((w for w in todas if w.title and "irys" in w.title.lower()), None)

    if not ventana_irys:
        logging.warning("⚠ No se encontró la ventana iRYS.")
        return None

    # Minimizar ventanas conflictivas
    for ventana in todas:
        if ventana.title and any(v.lower() in ventana.title.lower() for v in VENTANAS_MINIMIZAR):
            try:
                ventana.minimize()
            except:
                pass

    hwnd = ventana_irys._hWnd
    SW_RESTORE = 9
    SW_MAXIMIZE = 3
    SW_SHOW = 5

    try:
        ctypes.windll.user32.ShowWindow(hwnd, SW_RESTORE)
        time.sleep(0.3)
        ctypes.windll.user32.ShowWindow(hwnd, SW_MAXIMIZE)
        time.sleep(0.3)
        ctypes.windll.user32.ShowWindow(hwnd, SW_SHOW)
        ctypes.windll.user32.SetForegroundWindow(hwnd)
        ctypes.windll.user32.BringWindowToTop(hwnd)
        ctypes.windll.user32.SetActiveWindow(hwnd)

        logging.info("✅ iRYS restaurado, maximizado y al frente.")
    except Exception as e:
        logging.warning(f"No se pudo traer iRYS al frente: {e}")

    return ventana_irys


def esperar_y_clickear(imagen, timeout=15, confidence=0.85):
    """
    Espera a que aparezca la imagen en pantalla y hace clic en el centro.
    Si no aparece en 'timeout' segundos, devuelve False.
    """
    logging.info(f"Buscando: {imagen}")

    start = time.time()
    while time.time() - start < timeout:
        pos = pyautogui.locateCenterOnScreen(imagen, confidence=confidence)
        if pos:
            logging.info(f"✔ Encontrado {imagen} en {pos}, clic…")
            pyautogui.click(pos.x, pos.y, duration=0.4)
            time.sleep(0.8)
            return True
        time.sleep(0.5)

    logging.error(f"No se encontró {imagen} después de {timeout} s.")
    return False


def abrir_archivo_con_pyautogui(ruta_archivo):
    """Abre un archivo dentro de iRYS usando automatización GUI."""
    logging.info(f"Abriendo archivo: {ruta_archivo}")
    time.sleep(8)

    ventana_irys = maximizar_y_traer_irys()
    if not ventana_irys:
        return

    pyperclip.copy(ruta_archivo)
    time.sleep(5)

    # ===== SECUENCIA OBLIGATORIA =====
    if not esperar_y_clickear("captureFile.png"):
        return

    if not esperar_y_clickear("captureImport.png"):
        return

    if not esperar_y_clickear("captureImport3d.png"):
        return

    # 3) Intentar pegar con ctrl+v
    time.sleep(2)
    pyautogui.hotkey("ctrl", "v")
    pyautogui.press("enter")
    
    time.sleep(1)

    if not esperar_y_clickear("capturevolumetricData.png"):
        return

# ============================================================
#  MONITOREO DE CARPETAS
# ============================================================

class NuevoArchivoHandler(FileSystemEventHandler):
 def on_created(self, event):
        # Solo procesar si es un directorio dentro de carpeta_export
        if event.is_directory and os.path.dirname(event.src_path) == carpeta_export:
            carpeta_nueva = event.src_path
            logging.info(f"Nueva carpeta detectada: {os.path.basename(carpeta_nueva)}")

            primer_archivo = None
            timeout = 10
            t0 = time.time()

            # Esperar hasta que aparezca al menos un archivo
            while time.time() - t0 < timeout:
                archivos = [
                    f for f in os.listdir(carpeta_nueva)
                    if os.path.isfile(os.path.join(carpeta_nueva, f))
                ]
                if archivos:
                    primer_archivo = os.path.join(carpeta_nueva, archivos[0])
                    # Normalizar ruta para Windows
                    primer_archivo = os.path.normpath(primer_archivo)
                    logging.info(f"Primer archivo detectado: {primer_archivo}")
                    break
                time.sleep(0.5)

            if not primer_archivo:
                logging.warning("No se encontró archivo dentro de la carpeta nueva.")
                return

            # Lanza NNT e intenta abrir el archivo
            for intento in range(1, REINTENTOS + 1):
                ok, detalle = lanzar_nnt()
                if ok:
                    logging.info(f"[OK] NNT lanzado correctamente ({detalle})")
                    abrir_archivo_con_pyautogui(primer_archivo)
                    return

                logging.warning(f"Intento {intento} falló: {detalle}")
                if intento < REINTENTOS:
                    time.sleep(REINTENTO_DELAY)

            logging.error("No se pudo lanzar NNT tras todos los intentos.")

# ============================================================
#  MAIN
# ============================================================

if __name__ == "__main__":
    if not os.path.isdir(carpeta_export) or not os.path.isdir(folderirys):
        logging.error("Rutas en config.json NO válidas.")
        raise SystemExit(1)

    logging.info(f"Monitoreando carpeta: {carpeta_export}  (log en {LOG_FILE})")

    observer = Observer()
    observer.schedule(NuevoArchivoHandler(), carpeta_export, recursive=False)
    observer.start()

    try:
        while True:
            time.sleep(1)
    except KeyboardInterrupt:
        observer.stop()
        logging.info("Cerrando monitor...")

    observer.join()
