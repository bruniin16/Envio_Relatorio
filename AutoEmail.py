import pyautogui
import pyperclip
import pandas as pd
from time import sleep

pyautogui.PAUSE = 1

pyautogui.click(x=52, y=1064)
pyautogui.write("chrome")
pyautogui.press("enter")
sleep(8)

pyperclip.copy("https://drive.google.com/drive/folders/149xknr9JvrlEnhNWO49zPcw0PW5icxga")
pyautogui.hotkey("ctrl", "v")
pyautogui.press("enter")
sleep(3)

pyautogui.click(x=390, y=263, clicks=2)
sleep(2)
pyautogui.click(x=390, y=263)
pyautogui.click(x=1715, y=158)
pyautogui.click(x=1563, y=557)
sleep(5)
pyautogui.press("enter")

tabela = pd.read_excel(r".\Vendas - Dez.xlsx")
quantidade = tabela["Quantidade"].sum()
faturamento = tabela["Valor Final"].sum()

pyautogui.hotkey("ctrl", "t")
pyperclip.copy("https://mail.google.com/mail/u/0/#inbox")
pyautogui.hotkey("ctrl", "v")
pyautogui.press("enter")

sleep(3)

pyautogui.click(x=131, y=175)
pyautogui.write("EMAIL DO DESTINATÁRIO")
pyautogui.press("tab", presses=2, interval=0.5)
pyperclip.copy("Relatório de Vendas")
pyautogui.hotkey("ctrl", "v")
pyautogui.press("tab")

pyperclip.copy(f"""Bom dia,

Segue o relatório das vendas do dia anterior:

Faturamento total: R${faturamento:,.2f}
Quantidade total de itens vendidos: {quantidade:,}
""")

pyautogui.hotkey("ctrl", "v")
sleep(1)

pyautogui.hotkey("ctrl", "enter")