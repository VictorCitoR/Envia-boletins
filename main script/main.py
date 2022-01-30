import pandas as pd
import glob
import os
import PySimpleGUI as sg
from selenium import webdriver  # pip install selenium
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from webdriver_manager.chrome import ChromeDriverManager  # pip install webdriver_manager
import time

# Abre o Chrome
driver = webdriver.Chrome(ChromeDriverManager().install())
driver.get('https://web.whatsapp.com/')  # abre o site Whatsapp Web


# Contatos/Grupos - Informar o nome(s) de Grupos ou Contatos que serao enviadas as mensagens
def sem_erros(func):
    try:
        eval(func)
        return True
    except KeyError:
        return False


caminho = f'{os.path.abspath(os.getcwd())}\\ENVIO DE BOLETINS'

try:
    os.mkdir(caminho)
except FileExistsError:
    pass

excelFiles = glob.glob(f'{caminho}\\*.xlsm')
tabela = pd.read_excel(excelFiles[0],
                       sheet_name="ALUNOS")
for coluna in tabela:
    x = 0
    while sem_erros('tabela[coluna][x]'):
        x += 1

alunos = {}
for z in range(x):
    nome = tabela['NOME COMPLETO'][z]
    alunos[nome] = tabela['Telefone (Formato: 55dd99999999)'][z]


# Funcao que pesquisa o Contato/Grupo
def buscar_contato(tel):
    campo_pesquisa = driver.find_element(By.XPATH, '//div[contains(@class,"copyable-text selectable-text")]')
    time.sleep(1)
    campo_pesquisa.click()
    campo_pesquisa.send_keys(tel)
    campo_pesquisa.send_keys(Keys.ENTER)


# Funcao que envia a mensagem
def enviar_mensagem(mensagem, ns):
    campo_mensagem = driver.find_elements(By.XPATH, '//div[contains(@class,"copyable-text selectable-text")]')
    campo_mensagem[1].click()
    time.sleep(1)
    campo_mensagem[1].send_keys(f"Olá, {mensagem}, aqui está seu boletim do {ns}° simulado, "
                                f"compareça à coordenação para sanar qualquer eventual dúvida e obter orientações")
    campo_mensagem[1].send_keys(Keys.ENTER)


# Funcao que envia midia como mensagem
def enviar_midia(midia):
    driver.find_element(By.CSS_SELECTOR, "span[data-icon='clip']").click()
    driver.find_element(By.CSS_SELECTOR, "span[data-icon='clip']").click()
    attach = driver.find_element(By.CSS_SELECTOR, "input[type='file']")
    attach.send_keys(f"{caminho}\\BOLETIM_{midia}.jpg")
    time.sleep(1)
    send = driver.find_element(By.CSS_SELECTOR, "span[data-icon='send']")
    send.click()


layout = [
    [sg.Text('Quando obtiver acesso a sua conta do Whatsappp, indique o número\n'
             'do simulado e clique no "OK" abaixo para prosseguir.')],
    [sg.Text('Simulado N°'), sg.Input(key='NumeroDoSimulado')],
    [sg.Button('OK')]
]

janela = sg.Window('Confirmação', layout)
while True:
    event, values = janela.read()
    if event == sg.WIN_CLOSED or event == 'OK':  # if user closes window or clicks cancel
        break

janela.close()

# Percorre todos os contatos/Grupos e envia as mensagens
for aluno in alunos:
    telefone = str(alunos[aluno])
    nome = str(aluno)
    buscar_contato(telefone)
    enviar_mensagem(nome, values['NumeroDoSimulado'])
    enviar_midia(nome)
    time.sleep(1)
