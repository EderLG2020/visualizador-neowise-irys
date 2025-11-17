import cv2
import numpy as np
import mss
import time

def grabar_pantalla_optimizado(salida="grabacion.mp4", fps=3):
    with mss.mss() as sct:
        monitor = sct.monitors[1]

        ancho = monitor["width"]
        alto = monitor["height"]

        fourcc = cv2.VideoWriter_fourcc(*"mp4v")
        video = cv2.VideoWriter(salida, fourcc, fps, (ancho, alto))

        frame_interval = 1.0 / fps
        print(f"[OK] Grabando a {fps} FPS (Optimizado)")

        try:
            next_frame_time = time.time()

            while True:
                # Capturar frame raw (sin conversión)
                frame = sct.grab(monitor)

                # Convertir BGRA → BGR pero usando slicing (mucho más rápido)
                bgr = np.frombuffer(frame.rgb, dtype=np.uint8)
                bgr = bgr.reshape((alto, ancho, 3))

                video.write(bgr)

                next_frame_time += frame_interval
                sleep_time = next_frame_time - time.time()

                if sleep_time > 0:
                    time.sleep(sleep_time)

        except KeyboardInterrupt:
            print("\nGrabación detenida.")

        finally:
            video.release()
            print(f"Archivo guardado como: {salida}")

if __name__ == "__main__":
    grabar_pantalla_optimizado("grabacion.mp4", fps=5)
