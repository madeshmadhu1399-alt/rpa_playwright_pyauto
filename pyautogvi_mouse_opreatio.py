import pyautogui
import time

# mouse operation
pyautogui.moveTo(100, 100, duration=1)  # Move the mouse to (100, 100) over 1 second
pyautogui.click(100,100)  # Click the mouse at the current position
pyautogui.doubleClick(100, 100)  # Double-click the mouse at the current position
pyautogui.rightClick(100, 100)  # Right-click the mouse at the current position
pyautogui.scroll(10)  # Scroll up 10 units  
pyautogui.scroll(-10)  # Scroll down 10 units   
pyautogui.dragTo(200, 200, duration=1)  # Drag the mouse to (200, 200) over 1 second
pyautogui.dragRel(100, 0, duration=1)  # Drag the mouse 100 pixels to the right over 1 second