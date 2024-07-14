import win32com.clientclra
from datetime import datetime, timedelta, timezone
import os
import zipfile
from sys import exit
from datetime import date

import win32com.client as win32

arquivos_procurados = ["<lista_de_arquivos_procurados_+_extensao>"]
#Exemplo: "Nome_do_arquivo.xlsx"

outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")

#Direcionando o script para capturar as informações da caixa 6(inbox)
inbox = outlook.GetDefaultFolder(6)  

messages = inbox.Items
message = messages.GetLast()
body_content = message.body

#Capturar os arquivos do dia
received_dt = datetime.now() - timedelta(days=0) 
received_dt = received_dt.strftime('%d/%m/%Y 0:00am')
messages = messages.Restrict("[ReceivedTime] >= '" + received_dt + "'")

#Endereço de de e-mail que extrai os arquivos
messages = messages.Restrict("[SenderEmailAddress] = '<email_desejado>'")
outputDir = r"\\<diretorio_que_ira_salvar>"


contagem = 0
for messages in list(messages):

    s = messages.sender
    for attachment in messages.Attachments:
        # print(attachment)
        attachment.SaveASFile(os.path.join(outputDir, attachment.FileName))
        #print(f"anexo de salvo")
        if str(attachment) in arquivos_procurados:
            arquivos_procurados.remove(str(attachment))
        contagem += 1
        #Localizando o arquivo em ZIP e extraindo para CVS
        fantasy_zip = zipfile.ZipFile(r'\\<descompactar_arquivos_zip>')
        fantasy_zip.extractall(r'\\<diretorio_final>')

#Caso não localize os arquivos retorne "Não foram encontrados os seguintes arquivos"
if len(arquivos_procurados) > 0:
    print(f"Não foram encontrados os seguintes arquivos: \n {arquivos_procurados}")
else:
    #Caso localize todos os arquivos
    print(f"Procurando emails enviados a partir de '{received_dt}' \n")
    print(f"Arquivos baixados e disponibilizados em: \n {outputDir}")