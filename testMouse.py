import pyautogui
import time

print("🖱 Move your mouse to the File menu and wait...")
time.sleep(5)
x, y = pyautogui.position()
print(f"📍 File menu position: x={x}, y={y}")
