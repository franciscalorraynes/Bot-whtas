# importando as bibliotecas do selenium
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

# importando as bibliotecas do pyautogui
import pyautogui as tempoEspera
import pyautogui as teclasTeclados

# para trabalhar bibliotecas mais recentes
from selenium.webdriver.common.by import By

# para trabalhar com o excel
from openpyxl import load_workbook

# caminho do arquivo + o nome do arquivo no computador
name_file_contact = r"caminho para o arquivo que armazen os nomes dos contatos Contatos.xlsx"
sheet_data_contact = load_workbook(name_file_contact)

# selecionando o sheet de dados
sheet_select = sheet_data_contact["Dados"]

# emulando no navegador Chrome
browser = webdriver.Chrome()

# Passando e abrindo a pagina web
browser.get("https://web.whatsapp.com/")
# Espera até que o QR Code seja carregado na página (aqui esperamos 30 segundos)
# wait = WebDriverWait(browser, 30)
# wait.until(EC.presence_of_element_located((By.ID, 'side')))

# enquanto o tamanho da lista for menor 1, fica tentando fazer login a cada 3sec
while len(browser.find_elements(By.ID, "side")) < 1:
    # espera 3sec para ver se o whatsweb fez o login
    tempoEspera.sleep(3)

tempoEspera.sleep(3)

for i in range(2, len(sheet_select["A"]) + 1):
    # variaveis nome e mensagem
    contact_name = sheet_select["A%s" % i].value
    menssage_contact = sheet_select["B%s" % i].value

    # busca e descreve o nome via XPATH na barra de pesquisa do whatsapp
    browser.find_element(By.XPATH, '//*[@id="side"]/div[1]/div/div[2]/div[2]/div/div[1]/p').send_keys(contact_name)
    tempoEspera.sleep(3)

    teclasTeclados.press("enter")
    tempoEspera.sleep(3)

    browser.find_element(By.XPATH, '//*[@id="main"]/footer/div[1]/div/span[2]/div/div[2]/div[1]/div/div[1]/p').send_keys(menssage_contact)
    tempoEspera.sleep(3)
    
    teclasTeclados.press("enter")
    tempoEspera.sleep(3)

    
