import pyautogui
import pyperclip
import time
import pandas as pd

pyautogui.PAUSE = 1

pyautogui.press('win')
pyautogui.write('chrome')
pyautogui.press('enter')
pyperclip.copy('https://drive.google.com/drive/folders/1uileaevywoNZO5QtbEJyW8DDpcLk3SA7?usp=sharing')
pyautogui.hotkey('ctrl', 'v')
pyautogui.press('enter')
pyautogui.press('f11')

time.sleep(7)

pyautogui.click(x=379, y=236)
pyautogui.click(x=1155, y=90)
pyautogui.click(x=1029, y=432)

time.sleep(7)

table = pd.read_excel(r"C:\Users\Beatriz\Downloads\Vendas - Dez.xlsx")

quantity = table["Quantidade"].sum()
incoming = table["Valor Final"].sum()

pyautogui.press('f11')
pyautogui.hotkey("ctrl", "t")
pyperclip.copy('https://mail.google.com/mail/u/0/#inbox')
pyautogui.hotkey('ctrl', 'v')
pyautogui.press('enter')
pyautogui.press('f11')

time.sleep(7)

pyautogui.click(x=74, y=100)
time.sleep(2)

pyautogui.write("btourinhon@gmail.com")
pyautogui.press('enter')
pyautogui.press('tab')
pyautogui.write("Sales Report")
pyautogui.press('tab')


texto = f"""Dear bord of directors,

The incoming of the analyzed period was R${incoming:,.2f}
The amount of sold products was {quantity:,}

Cordially,
Beatriz Tourinho"""

pyautogui.write(texto)

pyautogui.hotkey('ctrl', 'enter')
pyautogui.press('f11')