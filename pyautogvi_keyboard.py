import pyautogui
import time

pyautogui.typewrite("Hello, World!" , interval=0.25)  # Type the string "Hello, World!"
pyautogui.press("enter")  # Press the Enter key
pyautogui.hotkey("cmd", "s")  # Press cmd+S to save
pyautogui.hscroll(10)  # Scroll horizontally to the right by 10 units
