import pyautogui
import time
import os

# Cria a pasta "prints" se não existir
os.makedirs("prints", exist_ok=True)

print("⏳ Você tem 5 segundos para focar a tela do SIGEAN...")
time.sleep(5)

for i in range(1, 9):  # ajuste o número total de capturas
    filename = f"prints/atendimento_{i:03}.png"
    screenshot = pyautogui.screenshot()
    screenshot.save(filename)
    print(f"📸 Print salvo: {filename}")

    pyautogui.press("f8")  # avança para a próxima triagem
    time.sleep(2)