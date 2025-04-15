import pyautogui
import time

# print("🖱️ Mova o mouse para o local desejado. Posso mostrar a posição por 10 segundos:")
# for i in range(10):
#     x, y = pyautogui.position()
#     print(f"Posição atual: x={x}, y={y}")
#     time.sleep(1)


print("⏳ Você tem 5 segundos para focar o navegador...")
time.sleep(5)

pyautogui.click(x=970, y=699)  # Atualize aqui
time.sleep(1)

pyautogui.typewrite("casa")
pyautogui.press("enter")
