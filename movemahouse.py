import pyautogui
import time

# Espera que se abra NNT.exe
time.sleep(6)

# Mueve el ratón a una coordenada (ajusta con pyautogui.position())
pyautogui.moveTo(8, 28, duration=0.5)
pyautogui.click()  # clic en "Archivo" por ejemplo

pyautogui.moveTo(70, 80, duration=0.5)
pyautogui.click()  # clic en "Abrir"

pyautogui.moveTo(309, 106, duration=0.5)
pyautogui.click()  # clic en "Abrir"
# Escribe ruta del archivo y presiona Enter
time.sleep(1)
pyautogui.write(r"D:\carpetalectura\paciente01") # aqui deberia cargar la ruta que pedi completo
pyautogui.press("enter")
