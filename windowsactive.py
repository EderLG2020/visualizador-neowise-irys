import pygetwindow as gw
import logging
import time
import ctypes

logging.basicConfig(level=logging.INFO, format="%(asctime)s [%(levelname)s] %(message)s")

# Lista de ventanas que queremos minimizar si aparecen
VENTANAS_MINIMIZAR = ["Neowise"]

def traer_al_frente_y_maximizar(hwnd):
    """Restaura, maximiza y trae al frente una ventana usando API de Windows."""
    SW_RESTORE = 9
    SW_SHOW = 5
    SW_MAXIMIZE = 3
    try:
        # Restaurar si está minimizada
        ctypes.windll.user32.ShowWindow(hwnd, SW_RESTORE)
        time.sleep(0.3)

        # Maximizar ventana
        ctypes.windll.user32.ShowWindow(hwnd, SW_MAXIMIZE)
        time.sleep(0.3)

        # Traer al frente
        ctypes.windll.user32.ShowWindow(hwnd, SW_SHOW)
        ctypes.windll.user32.SetForegroundWindow(hwnd)
        ctypes.windll.user32.BringWindowToTop(hwnd)
        ctypes.windll.user32.SetActiveWindow(hwnd)
        logging.info("✅ Ventana iRYS restaurada, maximizada y puesta al frente.")
        return True
    except Exception as e:
        logging.warning(f"No se pudo traer al frente o maximizar iRYS: {e}")
        return False

def ajustar_ventanas():
    """Minimiza todo menos iRYS, y deja iRYS como principal a pantalla completa."""
    todas_ventanas = gw.getAllWindows()
    ventana_irys = None

    for ventana in todas_ventanas:
        if not ventana.title:
            continue
        titulo = ventana.title.lower()

        # Si es iRYS -> guardamos referencia
        if "irys" in titulo:
            ventana_irys = ventana
            continue

        # Si coincide con alguna de las ventanas a minimizar
        if any(v.lower() in titulo for v in VENTANAS_MINIMIZAR):
            try:
                ventana.minimize()
                logging.info(f"Se minimizó la ventana: {ventana.title}")
            except Exception as e:
                logging.warning(f"No se pudo minimizar {ventana.title}: {e}")

    if ventana_irys:
        time.sleep(1)
        try:
            # Restaurar, maximizar y traer al frente
            ventana_irys.restore()
            time.sleep(0.5)
            ventana_irys.maximize()
            time.sleep(0.5)
            traer_al_frente_y_maximizar(ventana_irys._hWnd)
            logging.info(f"✅ iRYS en pantalla completa y como ventana activa: {ventana_irys.title}")
        except Exception as e:
            logging.warning(f"No se pudo restaurar/maximizar iRYS: {e}")
    else:
        logging.warning("⚠ No se encontró ninguna ventana de iRYS.")

if __name__ == "__main__":
    ajustar_ventanas()
