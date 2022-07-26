import pyautogui  # https://pyautogui.readthedocs.io/en/latest/
import pyperclip
import time
import os
import pandas as pd

# Configs
pyautogui.PAUSE = 2  # Delay entre as execuções
sistema_url = 'https://drive.google.com/drive/folders/149xknr9JvrlEnhNWO49zPcw0PW5icxga'
nome_arquivo = 'Vendas.xlsx'
pasta_download = os.path.join(os.environ['userprofile'], "Downloads", "Simulando")  # Pasta para salvar o arquivo
email_destinatario = 'emailaleatorio@maildrop.cc'


# Passo 1 - Abrindo o Google Chrome e entrando no "sistema":
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
time.sleep(1)  # Tempo para carregar
pyautogui.click(x=342, y=363)  # 1 click na planilha
time.sleep(1)  # Tempo para carregar
pyautogui.click(x=1739, y=178)  # click nos "..."
time.sleep(1)  # Tempo para carregar
pyautogui.click(x=1589, y=574)  # click em "Fazer download..."
time.sleep(5)  # Tempo para carregar
pyperclip.copy(nome_arquivo)  # Copiar nome do arquivo
time.sleep(1)  # Tempo para carregar
pyautogui.hotkey('ctrl', 'v')  # Colar nome do arquivo
time.sleep(1)  # Tempo para carregar
pyautogui.click(x=480, y=51)  # click na barra endereço
time.sleep(2)  # Tempo para carregar
pyperclip.copy(pasta_download)  # Copiar endereço da pasta onde salvar
pyautogui.hotkey('ctrl', 'v')
time.sleep(1)  # Tempo para carregar
pyautogui.press('enter')
time.sleep(1)  # Tempo para carregar
pyautogui.click(x=1124, y=681)  # click em "Salvar"
time.sleep(1)  # Tempo para carregar
pyautogui.click(x=1019, y=528)  # Confirmar substituição de arquivo já existente
time.sleep(5)  # Tempo para carregar

# Passo 3 - Gerar um relatório com as informações obtidas
caminho_arquivo = pasta_download + f'\{nome_arquivo}'  # Endereço completo do arquivo

tabela = pd.read_excel(rf'{caminho_arquivo}')
quantidade = tabela["Quantidade"].sum()  # Quantidade de itens vendidos
faturamento = tabela["Valor Final"].sum()  #Faturamento total

dia = time.strftime("%d/%m/%Y")


assunto_email = f'Relatório - {dia}'
conteudo_email = f"""
Prezados,

Relatório do dia {dia}:

O faturamento foi de: R${faturamento:,.2f}
A quantidade de produtos foi: R${quantidade:,.2f}

Grato.
"""

# Passo 4 - Entrar no e-mail e enviar

pyautogui.hotkey('ctrl', 't')  # Abrir nova aba
time.sleep(1)  # Tempo para carregar
pyperclip.copy('https://mail.google.com/mail/u/1/?ogbl#inbox?compose=new')  # Copiar url do gmail
time.sleep(1)  # Tempo para carregar
pyautogui.hotkey('ctrl', 'v')
time.sleep(1)  # Tempo para carregar
pyautogui.press('enter')  # Entrar no gmail em escrever e-mail
time.sleep(5)  # Tempo para carregar
pyperclip.copy(email_destinatario)  # E-mail de destinatário
time.sleep(1)  # Tempo para carregar
pyautogui.hotkey('ctrl', 'v')
time.sleep(1)  # Tempo para carregar
pyautogui.press('tab')
time.sleep(1)  # Tempo para carregar
pyperclip.copy(assunto_email)  # Assunto do e-mail
time.sleep(1)  # Tempo para carregar
pyautogui.hotkey('ctrl', 'v')
time.sleep(1)  # Tempo para carregar
pyautogui.press('tab')
time.sleep(1)  # Tempo para carregar
pyperclip.copy(conteudo_email)  # Conteúdo do e-mail
time.sleep(1)  # Tempo para carregar
pyautogui.hotkey('ctrl', 'v')
time.sleep(1)  # Tempo para carregar
pyautogui.hotkey('ctrl', 'enter')
