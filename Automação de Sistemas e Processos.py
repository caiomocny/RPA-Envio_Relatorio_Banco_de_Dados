#Passo 0: Importar as bibliotecas
import pyautogui
import pyperclip
from time import sleep

#Passo 1: Acessar o sistema
pyautogui.hotkey("ctrl", "t") #abrir uma nova aba com o atalho "ctrl + t"
pyperclip.copy("https://drive.google.com/drive/folders/149xknr9JvrlEnhNWO49zPcw0PW5icxga?usp=sharing") #copiar o link sem problema
#... com caracteres especiais
pyautogui.hotkey("ctrl", "v") #colar o link
sleep(1)
pyautogui.press("enter") # entrar no sistema

#Esperar o sistema carregar para poder interagir
sleep(5)

#Passo 2: Entrar na pasta correta
pyautogui.click(306, 269, clicks=2) #Entrar na pasta EXPORTAR
sleep(2)

#Passo 3: Baixar o arquivo
pyautogui.click(button="right") #clicar no arquivo
sleep(2)
pyautogui.click(398, 650) #baixar arquivo
sleep(5) #espera do download

# ### Vamos agora ler o arquivo baixado para pegar os indicadores de faturamento e quantidade de produtos.

#Passo 4: Importar o arquivo para o python
import pandas as pd

tabela = pd.read_excel(r"C:\Users\caiom\Downloads\Vendas - Dez.xlsx")

#Passo 5: Ler o arquivo e gerar os indicadores
faturamento = tabela["Valor Final"].sum()
display(tabela)
qntd = tabela['Quantidade'].sum()
display(faturamento, qntd)


# ### Vamos agora enviar um e-mail pelo gmail

#Passo 6: enviar arquivo
pyautogui.hotkey('ctrl', 't')
pyperclip.copy('https://mail.google.com/mail/u/1/#inbox') 
pyautogui.hotkey('ctrl', 'v')
pyautogui.press('enter') # Entrar no email
sleep(4)
pyautogui.click(x=106, y=214) # Clicar em Escrever email
sleep(2)
pyperclip.copy('camocny@gmail.com') 
pyautogui.hotkey('ctrl', 'v') # Escrever email e selecionar email 
pyautogui.press('enter') ## CERTIFICAR QUE NÃO TENHA ERRO NO EMAIL
pyautogui.press('tab')# Ir para o Assunto
pyperclip.copy("Automações.com -> Relatório de Vendas") # Escrever assunto
pyautogui.hotkey('ctrl', 'v')
pyautogui.press('tab') # Ir para o corpo do texto
pyautogui.write(f'''
Prezado,

O faturamento de ontem foi de R${faturamento:,.2f}
A quantidade de ontem vendida foi de {qntd:,}

Abs,
Caio Mocny
''') # Escrever corpo do texto
pyautogui.hotkey('ctrl', 'enter') # Encaminhar email
