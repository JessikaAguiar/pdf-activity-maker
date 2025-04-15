import pyautogui
import keyboard
import time

print("⏳ Você tem 5 segundos para focar o SIGEAN...")
time.sleep(5)

for i in range(900):  # Ajuste o número de páginas conforme necessário
    print(f"🖨️ Imprimindo página {i+1}...")

    pyautogui.press("f6")   # Imprime
    time.sleep(4)

    keyboard.press_and_release('f8')   # Próxima página
    time.sleep(2)

print("✅ Processo finalizado.")
