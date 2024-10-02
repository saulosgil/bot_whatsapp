"""
AUTOMATIZAÇÃO DE MENSAGENS DE WHATSAPP 
"""
import openpyxl
from urllib.parse import quote
import webbrowser
from time import sleep
import pyautogui
import os 

# Descomentar esse código para abrir o whatsapp web (quando não estiver logado)
# webbrowser.open('https://web.whatsapp.com/')
# sleep(30)

# Ler planilha e guardar informações sobre nome e telefone
workbook = openpyxl.load_workbook('names.xlsx')
pagina_clientes = workbook['Sheet1']

for linha in pagina_clientes.iter_rows(min_row=2):
    # nome e telefone
    nome = linha[0].value
    telefone = linha[1].value
    
    mensagem = f'Olá {nome} estou testando meu app de mensagens automatizadas.'
    
    # Criar links personalizados do whatsapp e enviar mensagens para cada pessoa
    # com base nos dados da planilha
    try:
        link_mensagem_whatsapp = f'https://web.whatsapp.com/send?phone={telefone}&text={quote(mensagem)}'
        webbrowser.open(link_mensagem_whatsapp)
        sleep(10)
        seta = pyautogui.locateCenterOnScreen('image_seta.png')
        sleep(10)
        pyautogui.click(seta[0], seta[1])
        sleep(10)
        pyautogui.hotkey('ctrl','w')
        sleep(10)
    except:
        print(f'Não foi possível enviar mensagem para {nome}')
        with open('erros.csv','a',newline='',encoding='utf-8') as arquivo:
            arquivo.write(f'{nome}, {telefone}{os.linesep}')