import pyautogui
import pyperclip
import time

pyautogui.PAUSE = 1

pyautogui.press('win')
pyautogui.write('chrome')
pyautogui.press('enter')

pyperclip.copy('https://drive.google.com/drive/folders/1uileaevywoNZO5QtbEJyW8DDpcLk3SA7?usp=sharing')
pyautogui.hotkey('ctrl', 'v')
pyautogui.press('enter')

pyautogui.hotkey('win', 'up')
time.sleep(5)
pyautogui.click(x=397, y=400)
pyautogui.click(x=1165, y=189)
pyautogui.click(x=987, y=526)
time.sleep(5)