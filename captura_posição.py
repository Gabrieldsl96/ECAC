import pyautogui

print("Mova o mouse sobre o botão e pressione Ctrl+C para capturar as coordenadas.")
try:
    while True:
        x, y = pyautogui.position()  # Obtém as coordenadas do mouse
        print(f"Posição atual do mouse: x={x}, y={y}", end="\r")  # Atualiza na mesma linha
except KeyboardInterrupt:
    print("\nCaptura de coordenadas interrompida.")
