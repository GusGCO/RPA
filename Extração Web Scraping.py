from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.action_chains import ActionChains
from selenium.common.exceptions import TimeoutException
from selenium.webdriver.support.ui import WebDriverWait
from datetime import datetime, timedelta, timezone
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from selenium import webdriver
import pandas as pd
import pyautogui
import logging
import shutil
import time
import csv
import os

# capturando o usuario logado na maquina
usuario_logado = os.environ.get("USERNAME").upper()

# criando dicionario com os usarios e infos para login no site
dic_login = {
    "785455":
    {'codigo':'697'
     ,'subcodigo':'111'
     ,'login1':'0001'
     ,'login2':'PHAB'
     ,'pass':'<senha>'
     ,'email':'gustavo.gco02@gmail.com'
     ,'final_doc':'456'
     ,'comeco_doc':'123'
    },

    "784848":
    {'codigo':'697'
     ,'subcodigo':'111'
     ,'login1':'0001'
     ,'login2':'GGCO'
     ,'pass':'<senha>'
     ,'email':'gustavo.gco02@gmail.com'
     ,'final_doc':'123'
     ,'comeco_doc':'456'},
     
}

# criando variavel com as infos do usuario logado
infos_usuario = dic_login[f'{usuario_logado}']

#Deletando o arquivo anterior com o mesmo nome

if os.path.exists(f"C:\\Users\\{usuario_logado}\\Downloads\\<nome_arquivo>.csv"):
    os.remove(f"C:\\Users\\{usuario_logado}\\Downloads\\<nome_arquivo>.csv")
    print("Arquivo deletado com sucesso!")
else:
    print("Arquivo não existe!")

    options = webdriver.ChromeOptions()

# colocando o chrome em segundo plano (comando abaixo)
#options.add_argument('--headless')

driver = webdriver.Chrome(executable_path="./driver/chromedriver.exe",options=options)
driver.get('<link_site>');

time.sleep(5)

#codigo
codigo = driver.find_element(By.XPATH,'/html/body/center/div/div[3]/form/div/table[1]/tbody/tr[1]/td[2]/input')
codigo.send_keys(infos_usuario['codigo'])

#subCodigo
subCod = driver.find_element(By.XPATH,'/html/body/center/div/div[3]/form/div/table[1]/tbody/tr[3]/td[2]/input')
subCod.send_keys(infos_usuario['subcodigo'])

#Login
Login = driver.find_element(By.XPATH,'/html/body/center/div/div[3]/form/div/table[1]/tbody/tr[5]/td[2]/input[1]')
Login.send_keys(infos_usuario['login1'])

Login1 = driver.find_element(By.XPATH,'/html/body/center/div/div[3]/form/div/table[1]/tbody/tr[5]/td[2]/input[2]')
Login1.send_keys(infos_usuario['login2'])

#senha
Login1 = driver.find_element(By.XPATH,'/html/body/center/div/div[3]/form/div/table[1]/tbody/tr[7]/td[2]/input')
Login1.send_keys(infos_usuario['pass'])


# aguardando 'x' segundos
time.sleep(5)

driver.find_element(By.XPATH,'//*[@id="btnLogin"]').click()

time.sleep(2)

posts  = driver.find_element(By.ID,"campos")

texto = posts.text

if texto == 'Informe seu e-mail:':
    # email = driver.find_elements('/html/body/div[8]/form/div[2]/div/div/input[1]')
    email = driver.find_element(By.XPATH,'/html/body/div[8]/form/div[2]/div/div/input[1]')
    email.send_keys(infos_usuario['email'])
elif texto == 'Complete seu CPF: XXX.XXX.XX':
    # cpf_u = driver.find_elements_element(By.XPATH'/html/body/div[8]/form/div[2]/div/div/input[1]')
    cpf_u = driver.find_element(By.ID,'validador')
    cpf_u.send_keys(Keys.COMMAND + (infos_usuario['final_doc']))
else:
    # cpf_p = driver.find_elements_element(By.XPATH'/html/body/div[8]/form/div[2]/div/div/input[1]')
    cpf_u = driver.find_element(By.XPATH,'/html/body/div[8]/form/div[2]/div/div/input[1]')
    cpf_u.send_keys((infos_usuario['comeco_doc']))

driver.find_element(By.XPATH,'//*[@id="campos"]/div/input[2]').click()
driver.find_element(By.XPATH,'/html/body/div[7]/div/div[2]/div[1]/div[1]/a').click()
time.sleep(10)

# mudando para a segunda aba que foi aberta
driver.switch_to.window(driver.window_handles[1])
time.sleep(5)

# encontrando o elemento que 'materializa' a tabela que vai ser extraida
driver.find_element(By.XPATH,'/html/body/div[2]/div/div/div[2]/div[2]/div/div/div/div/div[3]/div/button').click()

# aguardando 'x' segundos
time.sleep(30)

driver.find_element(By.ID,'searchResultExportLnk').click()
time.sleep(10)
driver.find_element(By.ID,'searchResultExportCSVLnk').click()
time.sleep(10)

time.sleep(90)

try:
    # reading two csv files
    df_historico = pd.read_csv(f'<diretorio_historico>.csv',keep_default_na=False)
    df_historico = df_historico.astype(str)
    
    df_atual = pd.read_csv(f"<diretorio_atual>.csv",keep_default_na=False)
    df_atual = df_atual.astype(str)
    
    df_final = pd.concat([df_historico, df_atual])
    df_final.to_csv(f"<diretorio_final>.csv",index=None)
# displaying result
# print(data3)
    
except OSError as e:
     print(e)
else:
    print("Processo finalizado!!")

len (df_historico) + len (df_atual)

len (df_atual)

pyautogui.alert("A execução finalizou!")