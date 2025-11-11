import os
import time
import subprocess
import ctypes
import logging
import pyautogui
from watchdog.observers import Observer
from watchdog.events import FileSystemEventHandler
import pygetwindow as gw  # Para minimizar la ventana de NNT

# ---------------- CONFIGURACIÓN ----------------
carpeta_export = r"D:\mundomedicodental\code\carpetalectura"
folderirys = r"D:\irys"
ruta_visor = os.path.join(folderirys, "NNT.exe")

ESPERA_USB_SEC = 6
REINTENTOS = 2
REINTENTO_DELAY = 3
LOG_FILE = "watch_nnt.log"
# ------------------------------------------------

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
    handlers=[
        logging.FileHandler(LOG_FILE, encoding="utf-8"),
        logging.StreamHandler()
    ]
)

def esperar_por_usb(timeout=ESPERA_USB_SEC):
    """Espera hasta que exista el archivo NNT.ini (dongle activo)."""
    start = time.time()
    while time.time() - start < timeout:
        if os.path.exists(os.path.join(folderirys, "NNT.ini")):
            return True
        time.sleep(0.5)
    return False

def shell_execute_runas(path, work_dir=None):
    """Ejecuta con permisos de administrador (UAC)."""
    SW_SHOWNORMAL = 1
    ret = ctypes.windll.shell32.ShellExecuteW(
        None, "runas", path, None, work_dir or os.path.dirname(path), SW_SHOWNORMAL
    )
    return ret > 32

def lanzar_nnt():
    """Intenta abrir NNT.exe con subprocess."""
    if not os.path.isfile(ruta_visor):
        return False, "No existe el ejecutable"

    if not esperar_por_usb():
        logging.warning("Timeout esperando driver/dongle. Intentando igual...")

    try:
        proc = subprocess.Popen([ruta_visor], cwd=folderirys)
        return True, f"popen (pid={proc.pid})"
    except Exception as e:
        logging.warning(f"Popen falló: {e}")
    return False, "Todos los métodos fallaron"

def abrir_archivo_con_pyautogui(ruta_archivo):
    """Usa PyAutoGUI para abrir un archivo en NNT."""
    logging.info(f"Abriendo archivo en NNT: {ruta_archivo}")

    time.sleep(6)  # Espera que NNT se abra completamente

    pyautogui.moveTo(8, 28, duration=0.5)
    pyautogui.click()  # Clic en "Archivo"

    pyautogui.moveTo(70, 80, duration=0.5)
    pyautogui.click()  # Clic en "Abrir"

    pyautogui.moveTo(309, 106, duration=0.5)
    pyautogui.click()  # Clic en "Abrir"

    time.sleep(1)
    pyautogui.write(ruta_archivo)
    pyautogui.press("enter")

class NuevoArchivoHandler(FileSystemEventHandler):
    def on_created(self, event):
        if not event.is_directory:
            return
        if os.path.dirname(event.src_path) != carpeta_export:
            return

        carpeta_nueva = event.src_path
        logging.info(f"Nueva carpeta detectada: {os.path.basename(carpeta_nueva)}")
        logging.info(f"Ruta completa de la carpeta: {carpeta_nueva}")

        # Esperar dinámicamente hasta que haya al menos un archivo en la carpeta
        primer_archivo = None
        timeout = 10  # segundos máximos de espera
        t0 = time.time()
        while time.time() - t0 < timeout:
            try:
                archivos = [f for f in os.listdir(carpeta_nueva)
                            if os.path.isfile(os.path.join(carpeta_nueva, f))]
                if archivos:
                    primer_archivo = os.path.join(carpeta_nueva, archivos[0])
                    logging.info(f"Primer archivo dentro de la carpeta: {primer_archivo}")
                    break
            except Exception as e:
                logging.warning(f"No se pudo leer el contenido de la carpeta: {e}")
            time.sleep(0.5)  # espera medio segundo antes de volver a intentar

        if not primer_archivo:
            logging.warning("No se encontró ningún archivo en la carpeta tras esperar.")
            return

       # ======= NUEVO: maximizar iRYS y minimizar otras ventanas de software de interés =======
        try:
            time.sleep(1)  # delay de 1 segundo antes de ejecutar

            # Obtener ventana activa
            ventana_activa = gw.getActiveWindow()
            if ventana_activa:
                logging.info(f"Ventana actualmente activa: {ventana_activa.title}")
            else:
                logging.info("No se pudo detectar la ventana activa.")

            todas_ventanas = gw.getAllWindows()

            # Maximizar iRYS primero
            time.sleep(10)  # espera 5 segundos después de maximizar
            ventana_irys = next((w for w in todas_ventanas if w.title and "iRYS" in w.title.lower()), None)
            if ventana_irys:
                ventana_irys.maximize()
                ventana_irys.activate()
                logging.info(f"Ventana iRYS maximizada y activada: {ventana_irys.title}")
            else:
                logging.info("No se encontró la ventana iRYS para maximizar.")

            # Minimizar Neowise u otras ventanas de interés
            ventanas_neowise = [w for w in todas_ventanas if w.title and "neowise" in w.title.lower()]
            for ventana in ventanas_neowise:
                ventana.minimize()
                logging.info(f"Se minimizó la ventana: {ventana.title}")

        except Exception as e:
            logging.warning(f"No se pudo ajustar ventanas de software de interés: {e}")

        # =======================================================

        # Lanzar NNT.exe
        for intento in range(1, REINTENTOS + 1):
            ok, detalle = lanzar_nnt()
            if ok:
                logging.info(f"[OK] NNT se ejecutó correctamente ({detalle})")
                abrir_archivo_con_pyautogui(primer_archivo)
                return
            logging.warning(f"Intento {intento} falló: {detalle}")
            if intento < REINTENTOS:
                time.sleep(REINTENTO_DELAY)

        logging.error("No se pudo lanzar NNT tras todos los reintentos.")

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







==============================================================================================================

import os
import time
import subprocess
import ctypes
import logging
import pyautogui
from watchdog.observers import Observer
from watchdog.events import FileSystemEventHandler
import pygetwindow as gw  # Para minimizar la ventana de NNT

# ---------------- CONFIGURACIÓN ----------------
carpeta_export = r"D:\mundomedicodental\code\carpetalectura"
folderirys = r"D:\irys"
ruta_visor = os.path.join(folderirys, "NNT.exe")

ESPERA_USB_SEC = 6
REINTENTOS = 2
REINTENTO_DELAY = 3
LOG_FILE = "watch_nnt.log"
# ------------------------------------------------

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
    handlers=[
        logging.FileHandler(LOG_FILE, encoding="utf-8"),
        logging.StreamHandler()
    ]
)

def esperar_por_usb(timeout=ESPERA_USB_SEC):
    """Espera hasta que exista el archivo NNT.ini (dongle activo)."""
    start = time.time()
    while time.time() - start < timeout:
        if os.path.exists(os.path.join(folderirys, "NNT.ini")):
            return True
        time.sleep(0.5)
    return False

def shell_execute_runas(path, work_dir=None):
    """Ejecuta con permisos de administrador (UAC)."""
    SW_SHOWNORMAL = 1
    ret = ctypes.windll.shell32.ShellExecuteW(
        None, "runas", path, None, work_dir or os.path.dirname(path), SW_SHOWNORMAL
    )
    return ret > 32

def lanzar_nnt():
    """Intenta abrir NNT.exe con subprocess."""
    if not os.path.isfile(ruta_visor):
        return False, "No existe el ejecutable"

    if not esperar_por_usb():
        logging.warning("Timeout esperando driver/dongle. Intentando igual...")

    try:
        proc = subprocess.Popen([ruta_visor], cwd=folderirys)
        return True, f"popen (pid={proc.pid})"
    except Exception as e:
        logging.warning(f"Popen falló: {e}")
    return False, "Todos los métodos fallaron"

def abrir_archivo_con_pyautogui(ruta_archivo):
    """Usa PyAutoGUI para abrir un archivo en NNT, asegurando que iRYS esté visible y al frente."""
    logging.info(f"Abriendo archivo en NNT: {ruta_archivo}")

    MAX_ESPERA = 90
    INTERVALO = 1.0
    SW_RESTORE = 9
    SW_SHOW = 5
    SW_MAXIMIZE = 3

    # === Esperar a que iRYS aparezca ===
    logging.info("Esperando a que iRYS se inicie y cree su ventana...")
    ventana_irys = None
    t0 = time.time()
    while time.time() - t0 < MAX_ESPERA:
        todas = gw.getAllWindows()
        ventana_irys = next((w for w in todas if w.title and "irys" in w.title.lower()), None)
        if ventana_irys:
            logging.info(f"Ventana iRYS detectada: {ventana_irys.title}")
            break
        time.sleep(INTERVALO)
    else:
        logging.error("❌ No se encontró la ventana iRYS tras esperar. Cancelando apertura.")
        return

    # === Esperar 6 segundos adicionales ===
    logging.info("Esperando 6 segundos adicionales para que iRYS termine de cargar...")
    time.sleep(6)

    # === Maximizar iRYS y minimizar otras ventanas ===
    try:
        todas_ventanas = gw.getAllWindows()
        # Minimizar Neowise u otros programas
        for ventana in todas_ventanas:
            if ventana.title and "neowise" in ventana.title.lower():
                ventana.minimize()
                logging.info(f"Se minimizó la ventana: {ventana.title}")

        # Maximizar iRYS correctamente
        hwnd = ventana_irys._hWnd
        ctypes.windll.user32.ShowWindow(hwnd, SW_RESTORE)
        time.sleep(0.3)
        ctypes.windll.user32.ShowWindow(hwnd, SW_MAXIMIZE)
        time.sleep(0.3)
        ctypes.windll.user32.ShowWindow(hwnd, SW_SHOW)
        ctypes.windll.user32.SetForegroundWindow(hwnd)
        ctypes.windll.user32.BringWindowToTop(hwnd)
        ctypes.windll.user32.SetActiveWindow(hwnd)
        logging.info("✅ iRYS restaurado, maximizado y puesto al frente.")
    except Exception as e:
        logging.warning(f"No se pudo ajustar ventanas: {e}")

    # === Confirmar foco activo ===
    t2 = time.time()
    confirmada = False
    while time.time() - t2 < 10:
        try:
            activa = gw.getActiveWindow()
            if activa and "irys" in activa.title.lower():
                confirmada = True
                logging.info(f"✅ Confirmado: iRYS es la ventana activa ({activa.title}).")
                break
        except Exception:
            pass
        time.sleep(0.5)

    if not confirmada:
        logging.warning("⚠ iRYS no llegó a ser la ventana activa. Se intentará continuar de todos modos.")

    # === Pequeña espera antes de los clics de PyAutoGUI ===
    time.sleep(2)

    # --- Clics para abrir el archivo ---
    pyautogui.moveTo(8, 28, duration=0.5)
    pyautogui.click()
    pyautogui.moveTo(70, 80, duration=0.5)
    pyautogui.click()
    pyautogui.moveTo(309, 106, duration=0.5)
    pyautogui.click()

    time.sleep(1)
    pyautogui.write(ruta_archivo)
    pyautogui.press("enter")

class NuevoArchivoHandler(FileSystemEventHandler):
    def on_created(self, event):
        if not event.is_directory:
            return
        if os.path.dirname(event.src_path) != carpeta_export:
            return

        carpeta_nueva = event.src_path
        logging.info(f"Nueva carpeta detectada: {os.path.basename(carpeta_nueva)}")
        logging.info(f"Ruta completa de la carpeta: {carpeta_nueva}")

        # Esperar dinámicamente hasta que haya al menos un archivo en la carpeta
        primer_archivo = None
        timeout = 10  # segundos máximos de espera
        t0 = time.time()
        while time.time() - t0 < timeout:
            try:
                archivos = [f for f in os.listdir(carpeta_nueva)
                            if os.path.isfile(os.path.join(carpeta_nueva, f))]
                if archivos:
                    primer_archivo = os.path.join(carpeta_nueva, archivos[0])
                    logging.info(f"Primer archivo dentro de la carpeta: {primer_archivo}")
                    break
            except Exception as e:
                logging.warning(f"No se pudo leer el contenido de la carpeta: {e}")
            time.sleep(0.5)  # espera medio segundo antes de volver a intentar

        if not primer_archivo:
            logging.warning("No se encontró ningún archivo en la carpeta tras esperar.")
            return

       # ======= NUEVO: maximizar iRYS y minimizar otras ventanas de software de interés =======
        try:
            time.sleep(1)  # delay de 1 segundo antes de ejecutar

            # Obtener ventana activa
            ventana_activa = gw.getActiveWindow()
            if ventana_activa:
                logging.info(f"Ventana actualmente activa: {ventana_activa.title}")
            else:
                logging.info("No se pudo detectar la ventana activa.")

            todas_ventanas = gw.getAllWindows()

            # Maximizar iRYS primero
            time.sleep(10)  # espera 5 segundos después de maximizar
            ventana_irys = next((w for w in todas_ventanas if w.title and "iRYS" in w.title.lower()), None)
            if ventana_irys:
                ventana_irys.maximize()
                ventana_irys.activate()
                logging.info(f"Ventana iRYS maximizada y activada: {ventana_irys.title}")
            else:
                logging.info("No se encontró la ventana iRYS para maximizar.")

            # Minimizar Neowise u otras ventanas de interés
            ventanas_neowise = [w for w in todas_ventanas if w.title and "neowise" in w.title.lower()]
            for ventana in ventanas_neowise:
                ventana.minimize()
                logging.info(f"Se minimizó la ventana: {ventana.title}")

        except Exception as e:
            logging.warning(f"No se pudo ajustar ventanas de software de interés: {e}")

        # =======================================================

        # Lanzar NNT.exe
        for intento in range(1, REINTENTOS + 1):
            ok, detalle = lanzar_nnt()
            if ok:
                logging.info(f"[OK] NNT se ejecutó correctamente ({detalle})")
                abrir_archivo_con_pyautogui(primer_archivo)
                return
            logging.warning(f"Intento {intento} falló: {detalle}")
            if intento < REINTENTOS:
                time.sleep(REINTENTO_DELAY)

        logging.error("No se pudo lanzar NNT tras todos los reintentos.")

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

