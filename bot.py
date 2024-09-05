import openpyxl
from urllib.parse import quote
import webbrowser
from time import sleep
import pyautogui

#Lendo a planilha e guardando informações sobre nome, telefone e data de vencimento

workbook = openpyxl.load_workbook('telefones.xlsx')
pagina_clientes = workbook['Planilha1']

for linha in pagina_clientes.iter_rows(min_row=2):

    nome = linha[0].value
    telefone = linha[1].value
    vencimento = linha[2].value

#Criando links personalizados do whatsapp e enviar mensagens para cada cliente
# https://web.whatsapp.com/send?phone=&text

    mensagem = f'Olá {nome} seu boleto vence no dia {vencimento.strftime('%d/%m/%Y')}. Favor pagar no link https://www.link_exemplo.com'

    try:

        link_mensagem_whatsapp = f'https://web.whatsapp.com/send?phone={telefone}&text={quote(mensagem)}'
        webbrowser.open(link_mensagem_whatsapp)
        sleep(12)
    
        seta = pyautogui.locateCenterOnScreen('Screenshot_1.png')
        sleep(5)
        pyautogui.click(seta[0],seta[1])
        sleep(5)
        pyautogui.hotkey('ctrl','w')
        sleep(5)

    except:

        print(f'Não foi possível enviar mensagem para {nome}')
        with open('erros.csv','a', newline='', encoding='utf-8') as arquivo:
            arquivo.write(f'{nome}, {telefone}')





