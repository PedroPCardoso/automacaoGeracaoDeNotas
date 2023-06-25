from selenium.webdriver.chrome.service import Service
from selenium import webdriver
from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver.common.by import By
import time
from selenium.webdriver.common.action_chains import ActionChains
from selenium.common.exceptions import NoSuchElementException
import os
import shutil
#CAPTCHAR
import pytesseract
import requests
from PIL import Image
from io import BytesIO
from datetime import date, timedelta
from calendar import monthrange

def makeLogin():
    driver.get("https://fisco.net.br/isseletronico/conceicaodojacuipe/")
    html = driver.page_source
    # print(html)
    iframe = driver.find_element(By.NAME, "mainframe")
    driver.switch_to.frame(iframe)

    driver.find_element(By.NAME, "txtLogin").send_keys(nome)
    driver.find_element(By.NAME, "txtSenha").send_keys(codigo)

    time.sleep(10)
def goListNotes():
    try:
        myDiv = driver.find_element(By.ID,"btnNotaFiscal")
        myDiv.click();
        iframe = driver.find_element(By.NAME, "conteudo")
        driver.switch_to.frame(iframe)
        myDiv = driver.find_element(By.ID,"btnCopiar")
        myDiv.click();
        iframe = driver.find_element(By.NAME, "direita")
        driver.switch_to.frame(iframe)
    except NoSuchElementException:
        time.sleep(5)
        goListNotes()
def downloadNotes(elements):
    for i, element in enumerate(elements):
        action = ActionChains(driver)
        action.move_to_element(element).click().perform()
        time.sleep(5)
        # Obtenha o nome do arquivo baixado
        nome_arquivo = max(
            [os.path.join(diretorio_download, f) for f in os.listdir(diretorio_download)],
            key=os.path.getctime,
        )
        # Copie o arquivo baixado para a pasta da empresa
        shutil.move(nome_arquivo, os.path.join(pasta_empresa, f"nota{i+1}.xml"))
#SCRIPT PARA O ARQUIVO

# Abrir o arquivo de texto
with open('empresas.txt', 'r') as file:
    # Ler as linhas do arquivo
    lines = file.readlines()

# Criar uma lista para armazenar as informações das empresas
empresas = []

# Variáveis temporárias para armazenar as informações de cada empresa
nome = ''
codigo = ''
empresa = ''

# Iterar pelas linhas do arquivo
for line in lines:
    # Remover espaços em branco e quebras de linha
    line = line.strip()

    # Verificar se a linha não está vazia
    if line:
        # Adicionar a informação à variável correspondente
        if not nome:
            nome = line
        elif not codigo:
            codigo = line
        elif not empresa:
            empresa = line

        # Verificar se todas as informações da empresa foram obtidas
        if nome and codigo and empresa:
            # Adicionar as informações à lista de empresas
            empresas.append((nome, codigo, empresa))

            # Limpar as variáveis temporárias para a próxima empresa
            nome = ''
            codigo = ''
            empresa = ''
# Iterar sobre as empresas e realizar as ações desejadas
for nome, codigo, empresa in empresas:
    i=1
    driver = webdriver.Chrome(executable_path="/usr/local/bin/chromedriver")
    makeLogin()
    nome_empresa = empresa    
    #caminho da empresa
    diretorio_download = "/home/pedro/Downloads"
    goListNotes()
    # Crie uma pasta com o nome da empresa no diretório de download
    pasta_empresa = os.path.join(diretorio_download, nome_empresa)
    os.makedirs(pasta_empresa, exist_ok=True)

    # Selecionar o checkbox com ID "chkListarEmitidas"
    # chkListarEmitidas_element =  driver.find_element(By.ID,'chkListarEmitidas')
    # chkListarEmitidas_element.click()


    # Obter a data atual
    data_atual = date.today()
    # Obter o primeiro dia do mês atual
    primeiro_dia_mes = date(data_atual.year, data_atual.month, 1)

    data_formatada = primeiro_dia_mes.strftime("%d/%m/%Y")
    # Localizar o elemento pelo ID "data1" e inserir a data formatada
    elemento_data = driver.find_element(By.ID,"data1")
    elemento_data.send_keys(data_formatada)

    data_atual = date.today()
    ultimo_dia_mes = monthrange(data_atual.year, data_atual.month)[1]
    data_final_mes = date(data_atual.year, data_atual.month, ultimo_dia_mes)
    data_formatada_final_mes = data_final_mes.strftime("%d/%m/%Y")

    # Localizar o elemento pelo ID "data2" e inserir a data formatada
    elemento_data_final =driver.find_element(By.ID,"data2")
    elemento_data_final.send_keys(data_formatada_final_mes)


    # Localizar o botão pelo ID "btnPesquisar" e clicar nele
    botao_pesquisar = driver.find_element(By.ID,"btnPesquisar")
    botao_pesquisar.click()

    time.sleep(5)

    # xpath = '//a[contains(@class, "fa-file-excel")]'
    elements = driver.find_elements(By.CLASS_NAME,"fas.fa-file-excel")
    # downloadNotes(elements)
    try:
        i=i+1;
        element = driver.find_element(By.LINK_TEXT,str(i))
        element.click()
        elements = driver.find_elements(By.CLASS_NAME,"fas.fa-file-excel")
        downloadNotes(elements)        
    except NoSuchElementException:
        print("ERROR");
    # Feche o navegador
    driver.quit()
