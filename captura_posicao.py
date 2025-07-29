import pyautogui
import time

print("Posicione o mouse sobre o botão 'Salvar Nota' em 10 segundos...")
time.sleep(10)

x, y = pyautogui.position()
print(f"Posição capturada: x={x}, y={y}")
