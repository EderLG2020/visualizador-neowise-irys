import pyautogui
import mouse

print("🖱️ Haz 3 clics en los lugares que quieras capturar (esperando...)")

coords = []

# Captura exactamente 4 clics
for i in range(1, 8):
    mouse.wait(button='left', target_types=('down',))
    pos = pyautogui.position()
    coords.append(pos)
    print(f"✅ Click {i}: {pos}")

# Mostrar resultados finales
print("\n📍 Coordenadas finales:")
for i, (x, y) in enumerate(coords, start=1):
    print(f"Click {i}: ({x}, {y})")

# Guardar en archivo
with open("click_positions.txt", "w") as f:
    for (x, y) in coords:
        f.write(f"{x},{y}\n")

print("\n💾 Coordenadas guardadas en 'click_positions.txt'")
