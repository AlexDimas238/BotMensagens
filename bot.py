import openpyxl
from urllib.parse import quote
import webbrowser
from time import sleep
import pyautogui

'''
webbrowser.open('https://web.whatsapp.com/')
sleep(10)
'''

# Ler Planilha e Guardar informações - Nome e Telefone
workbook = openpyxl.load_workbook('ClientesEX.xlsx')
pagina_clientes = workbook['Planilha1']

for linha in pagina_clientes.iter_rows(min_row=2):
    # Qual coluna está o nome e o telefone colocar entre colchetes
    nome = linha[1].value
    telefone = linha[3].value

    # Mensagem que vai ser enviada ao cliente
    mensagem = f'Ola {nome} Clique no link para entrar no grupo do nosso passeio https://chat.whatsapp.com/FEWtnLfpWwYERTjWuU5rb1'

    try:
        # Criar links personalizados do whatsapp
        link_msg_whatsapp = f'https://web.whatsapp.com/send?phone={telefone}&text={quote(mensagem)}'
        # Vai Abrir o navegador
        webbrowser.open(link_msg_whatsapp)
        sleep(10)
        # Varrer a tela até encontrar a imagem de enviar
        seta = pyautogui.locateCenterOnScreen('seta.png')
        sleep(15)
        pyautogui.click(seta[0], seta[1])
        sleep(5)
        # Fechar a Aba
        pyautogui.hotkey('ctrl', 'w')
        sleep(5)
    except:
        print(f'Não foi possível enviar mensagem para {nome}')
        with open('erros.csv', 'a', newline='', encoding='utf-8') as arquivo:
            arquivo.write(f'{nome, telefone}')
        pyautogui.hotkey('ctrl', 'w')
