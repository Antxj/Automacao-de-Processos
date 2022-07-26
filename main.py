import pyautogui  # https://pyautogui.readthedocs.io/en/latest/
import pyperclip
import time
import os

pyautogui.PAUSE = 2  # Delay entre as execuções

# Passo 1 - Abrindo o Google Chrome e entrando no "sistema":
sistema_url = 'https://drive.google.com/drive/folders/149xknr9JvrlEnhNWO49zPcw0PW5icxga'
pyautogui.press('win')
pyautogui.write('chrome')
time.sleep(1)  # Tempo para carregar
pyautogui.press('enter')
time.sleep(1)  # Tempo para carregar
pyautogui.hotkey('ctrl', 't')
pyperclip.copy(sistema_url)
pyautogui.hotkey('ctrl', 'v')
pyautogui.press('enter')
time.sleep(1)  # Tempo para o site carregar

# Passo 2 - Baixar a "base de dados" na pasta Exportar do "sistema".
pyautogui.click(x=340, y=276, clicks=2)
pyautogui.click(x=342, y=363)  # 1 click na planilha
pyautogui.click(x=1739, y=178)  # click nos "..."
pyautogui.click(x=1589, y=574)  # click em "Fazer download..."
time.sleep(5)  # Tempo para carregar
pyautogui.click(x=480, y=51)  # click na barra endereço
time.sleep(2)  # Tempo para carregar
pasta_download = os.path.join(os.environ['userprofile'], "Downloads", "Simulando")  # Pasta download padrão do Windows
pyperclip.copy(pasta_download)
pyautogui.hotkey('ctrl', 'v')
pyautogui.press('enter')
time.sleep(1)  # Tempo para carregar
pyautogui.click(x=1124, y=681)  # click em "Salvar"
time.sleep(3)  # Tempo para carregar

# Passo 3 - Gerar um relatório.

