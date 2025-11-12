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

# ---------------- CONFIGURACIÓN ----------------
carpeta_export = r"D:\mundomedicodental\code\carpetalectura"
folderirys = r"D:\irys"
ruta_visor = os.path.join(folderirys, "NNT.exe")

ESPERA_USB_SEC = 6
REINTENTOS = 2
REINTENTO_DELAY = 3
LOG_FILE = "watch_nnt.log"
VENTANAS_MINIMIZAR = ["Neowise"]
# ------------------------------------------------

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
    handlers=[
        logging.FileHandler(LOG_FILE, encoding="utf-8"),
        logging.StreamHandler()
    ]
)

# -----------------------------------------
# FUNCIONES DE UTILIDAD
# -----------------------------------------
def esperar_por_usb(timeout=ESPERA_USB_SEC):
    start = time.time()
    while time.time() - start < timeout:
        if os.path.exists(os.path.join(folderirys, "NNT.ini")):
            return True
        time.sleep(0.5)
    return False

def lanzar_nnt():
    if not os.path.isfile(ruta_visor):
        return False, "No existe el ejecutable"

    if not esperar_por_usb():
        logging.warning("Timeout esperando driver/dongle. Intentando igual...")

    try:
        proc = subprocess.Popen([ruta_visor], cwd=folderirys)
        logging.info(f"NNT.exe lanzado (pid={proc.pid}), esperando 8 segundos para que cargue...")
        time.sleep(8)  # <-- Delay desde que se abre el software
        return True, f"popen (pid={proc.pid})"
    except Exception as e:
        logging.warning(f"Popen falló: {e}")
    return False, "Todos los métodos fallaron"


def maximizar_y_traer_irys():
    """Maximiza iRYS, lo pone al frente y minimiza otras ventanas de interés."""
    todas = gw.getAllWindows()
    ventana_irys = next((w for w in todas if w.title and "irys" in w.title.lower()), None)

    if not ventana_irys:
        logging.warning("⚠ No se encontró la ventana iRYS.")
        return None

    # Minimizar otras ventanas
    for ventana in todas:
        if ventana.title and any(v.lower() in ventana.title.lower() for v in VENTANAS_MINIMIZAR):
            try:
                ventana.minimize()
                logging.info(f"Se minimizó la ventana: {ventana.title}")
            except:
                pass

    # Maximizar iRYS y traer al frente
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
        # Clic para asegurar foco
        # rect = ventana_irys.box
        # pyautogui.click(rect.left + 20, rect.top + 20)
        logging.info("✅ iRYS restaurado, maximizado y al frente.")
    except Exception as e:
        logging.warning(f"No se pudo maximizar/traer iRYS al frente: {e}")

    return ventana_irys

def abrir_archivo_con_pyautogui(ruta_archivo):
    logging.info(f"Abriendo archivo: {ruta_archivo}")

    # === Esperar 8 segundos antes de manipular iRYS ===
    logging.info("Esperando 8 segundos para que iRYS termine de cargarse...")
    time.sleep(8)

    ventana_irys = maximizar_y_traer_irys()
    if not ventana_irys:
        return
    pyperclip.copy(ruta_archivo)

    time.sleep(2)
    # Clics para abrir archivo
    pyautogui.click(15, 28, duration=0.5)
    pyautogui.click(83, 80, duration=0.5)
    pyautogui.click(331, 106, duration=0.5)
    time.sleep(1)
    pyautogui.hotkey("ctrl", "v")
    pyautogui.press("enter")
    time.sleep(1)
    pyautogui.click(652, 384, duration=0.5)
# -----------------------------------------
# MONITOREO DE CARPETAS
# -----------------------------------------
class NuevoArchivoHandler(FileSystemEventHandler):
    def on_created(self, event):
        if not event.is_directory or os.path.dirname(event.src_path) != carpeta_export:
            return

        carpeta_nueva = event.src_path
        logging.info(f"Nueva carpeta detectada: {os.path.basename(carpeta_nueva)}")

        # Esperar primer archivo
        primer_archivo = None
        timeout = 10
        t0 = time.time()
        while time.time() - t0 < timeout:
            archivos = [f for f in os.listdir(carpeta_nueva)
                        if os.path.isfile(os.path.join(carpeta_nueva, f))]
            if archivos:
                primer_archivo = os.path.join(carpeta_nueva, archivos[0])

                logging.info(f"Primer archivo detectado ignorando caracteres especiales : {primer_archivo}")
                break
            time.sleep(0.5)

        if not primer_archivo:
            logging.warning("No se encontró ningún archivo en la carpeta tras esperar.")
            return

        # Lanzar NNT y abrir archivo
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

# -----------------------------------------
# MAIN
# -----------------------------------------
if __name__ == "__main__":
    if not os.path.isdir(carpeta_export) or not os.path.isdir(folderirys):
        logging.error("Rutas configuradas no válidas.")
        raise SystemExit(1)

    logging.info(f"Monitoreando carpeta: {carpeta_export} ... (log en {LOG_FILE})")
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
