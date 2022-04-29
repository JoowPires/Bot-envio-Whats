from re import T
import time
from webdriver_manager.chrome import ChromeDriverManager
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
import pandas as pd
import urllib

contatos_df = pd.read_excel('Enviar.xlsx')

# Abrir navegador e setar opções do Chrome


def launchBrowser():
    chrome_options = Options()
    chrome_options.add_argument('start-maximized')
    chrome_options.add_experimental_option(
        'excludeSwitches', ['enable-logging'])
    global wd
    wd = webdriver.Chrome(options=chrome_options)
    wd.get("https://web.whatsapp.com")

# Capturando o QRCode na página


def locateQrCodeElem():
    time.sleep(5)
    try:
        qr_code = wd.find_element(
        By.XPATH, '//*[@id="app"]/div/div/div[2]/div[1]/div/div[2]/div/canvas')
        qr_code.is_displayed()
        input('Scan QRCode e pressione Enter')
    except:
        print("QRCode não encontrado")
        closeBrowser()


# Passando por cada contato

def checkEachContact():
    for i, mensagem in enumerate(contatos_df['Mensagem']):
        pessoa = contatos_df[i, "Pessoa"]
        numero = contatos_df[i, "Número"]
        texto = urllib.parse.quote(f"Olá Colaborador, {mensagem}")
        global link 
        link = f"https://web.whatsapp.com/send?phone={numero}&text={texto}"
        print(
            f'Enviando para mensagem para {pessoa}, número {numero}')
        sendMessage(pessoa, numero, mensagem)
# Enviar mensagem


def sendMessage(pessoa, numero, mensagem):
    print(f'Enviando mensagem para {pessoa}')
    wd.get(link)
    time.sleep(30)
    try: 
        (wd.find_element(By.XPATH, '//*[@id="main"]/footer/div[1]/div/span[2]/div/div[2]/div[2]').size != 0)
        time.sleep(5)
        print('Numero encontrado')
        sendBtn = wd.find_element(By.XPATH, '//*[@id="main"]/footer/div[1]/div/span[2]/div/div[2]/div[2]')
        sendBtn.click()
        print("Mensagem enviada!")
        time.sleep(10)
             
    except:
        print('Numero não encontrado')
        errorBtn = wd.find_element(By.XPATH, '//*[@id="app"]/div/span[2]/div/span/div/div/div/div/div/div[2]/div/div')
        errorBtn.click()
        time.sleep(5)

# Encerrar navegador


def closeBrowser():
    time.sleep(1)
    wd.close()

# Rodar o programa


def runScript():
    launchBrowser()
    locateQrCodeElem()
    checkEachContact()
    closeBrowser()


runScript()