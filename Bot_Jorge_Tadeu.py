#!/usr/bin/env python
# coding: utf-8

# In[ ]:


from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from __future__ import print_function
import pyautogui
import pymsgbox
import time
import datetime
import win32com.client as win32
import os.path
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError
import schedule 
import threading
from webdriver_manager.chrome import ChromeDriverManager
import telebot
import requests
import time
import json
import os

SCOPES = ['https://www.googleapis.com/auth/spreadsheets']

SAMPLE_SPREADSHEET_ID = 'SEU ID'
SAMPLE_RANGE_NAME = 'Chat Bot!A1:D8'

CHAVE_API = "SUA CHAVE API"
bot = telebot.TeleBot(CHAVE_API)

erros = []
list_err_chatbot = []

pymsgbox.alert('O bot irá começar','Jorge Tadeu')
navegador = webdriver.Chrome(ChromeDriverManager().install())
time.sleep(5)
pyautogui.hotkey('win','up')

#Chat Bot
def teste_elem(elemento, indice):
    if indice == 1:
        try:
            navegador.find_element_by_xpath(elemento).send_keys('NOME')
        except:
            erros.append(indice)
    elif indice == 2:
        try:
            navegador.find_element_by_xpath(elemento).send_keys('CPF')
        except:
            erros.append(indice)
    else:  
        try:
            navegador.find_element_by_xpath(elemento).click()
        except:
            erros.append(indice)

loc_elem = ['//*[@id="omni-show-chat"]', 
            '//*[@id="userNameForm"]', 
            '//*[@id="cpf"]',
            '//*[@id="omni-chat-web"]/div/div/div/div[2]/div[3]/button', 
            '//*[@id="chat-list-messages"]/li[1]/div/div[2]/div[3]/div/button[1]', 
            '//*[@id="chat-list-messages"]/li[3]/div/div[2]/div[3]/div/button[6]', 
            '//*[@id="chat-list-messages"]/li[5]/div/div[2]/div[3]/div/button[1]', 
            '//*[@id="chat-list-messages"]/li[7]/div/div[2]/div[3]/div/button[2]',
            '//*[@id="chat-list-messages"]/li[9]/div/div[2]/div[3]/div/button[1]']

def chatbot():
    creds = None
    if os.path.exists('token.json'):
        creds = Credentials.from_authorized_user_file('token.json', SCOPES)
        print(f'fora creds.valid: {creds.valid}')
    if not creds or not creds.valid:
        print(f'dentro do if creds.valid: {creds.valid}')
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(
                'client_secret.json', SCOPES)
            creds = flow.run_local_server(port=0)
        with open('token.json', 'w') as token:
            token.write(creds.to_json())

    try:
        service = build('sheets', 'v4', credentials=creds)

        sheet = service.spreadsheets()
        result = sheet.values().get(spreadsheetId=SAMPLE_SPREADSHEET_ID,
            range=SAMPLE_RANGE_NAME).execute()
        values = result.get('values', [])
        
#adicionar na planilha
        valores_adicionar = [list_err_chatbot]
        result = sheet.values().update(spreadsheetId='DADOSID',
                                    range=f'Chat Bot!C{indice}', valueInputOption="USER_ENTERED",
                                        body={"values": valores_adicionar}).execute()
    
        valores_adicionar1 = [
            ["ERRO", "=NOW()"]
        ]
        
        global i
        result = sheet.values().update(spreadsheetId=SAMPLE_SPREADSHEET_ID,
                                    range=f'Relatórios!A{i+1}', valueInputOption="USER_ENTERED",
                                        body={"values": valores_adicionar1}).execute()
        print("Adicionado com sucesso")
    except HttpError as err:
            print(err)
            
def func_chat_bot():
    dic_chat_bot = {
        2 : "Erro ao acessar o endereço https://meu.inss.gov.br/#/login",
        3 : "Erro ao localizar o avatar da assistente virtual, Helô, no canto inferior direito:",
        4 : "Erro ao informar seu nome e CPF, depois clique em Iniciar Atendimento:",
        5 : "Erro ao selecionar a opção contida na imagem abaixo:",
        6 : "Erro ao testar se o transbordo para o humano está funcionando, selecione os submenus tela a tela até que apareça a opção conseguiu encontrar o que procurava?",
        7 : "Selecione Não e depois selecione Sim para a pergunta “Você quer conversar com de nossos atendentes?”",
        8 : "Ao responder Sim, o esperado é que seja direcionado para a fila para ser atendido."
    }
    for i in range(len(erros)):
        err_chat_bot = dic_chat_bot[erros[i]]
        list_err_chatbot.append(err_chat_bot)
        print(list_err_chatbot)
        
def func_email() :
    outlook = win32.Dispatch('outlook.application')
    email = outlook.CreateItem(0)

    # configurar as informações do seu e-mail
    email.To = "EMAIL"
    email.Subject = "Erros"
    email.HTMLBody = f"""
    <p>Olá eu sou o Jorge Tadeu</p>

    <p>Chat bot: </p>
    <p>{list_err_chatbot}</p>

    """
    
    email.Send()
    print("Email Enviado")
    
def func_screen_chatbot():
    navegador.get_screenshot_as_file(f"{erros}.png")
    time.sleep(5)
    print("printei")
    
@bot.message_handler(commands=["Erros_Diarios"])
def Erros_Diarios(mensagem):
    if list_err_chatbot == []:
        bot.send_message(mensagem.chat.id, "Hoje o Chat Bot não apresentou nenhum erro")
    else:
        bot.send_message(mensagem.chat.id, f"{list_err_chatbot}")

@bot.message_handler(commands=["Erros_Mensais"])
def Erros_Mensais(mensagem):
    bot.send_message(mensagem.chat.id, "Relatórios mensais")

@bot.message_handler(commands=["Menu_inicial"])
def Menu_inicial(mensagem):
    verificar()

@bot.message_handler(commands=["Chat_Bot"])
def Chat_Bot(mensagem):
    texto = """
    O que você deseja (Clique em uma opção):
    /Erros_Diarios
    /Erros_Mensais
    /Menu_inicial"""
    bot.send_message(mensagem.chat.id, texto)

@bot.message_handler(commands=["SAT_Central"])
def SAT_Central(mensagem):
    bot.send_message(mensagem.chat.id, "Para enviar uma reclamação, mande um e-mail para reclamação@balbalba.com")

@bot.message_handler(commands=["Meu_INSS"])
def Meu_INSS(mensagem):
    bot.send_message(mensagem.chat.id, "Valeu! Lira mandou um abraço de volta")

def verificar(mensagem):
    return True

@bot.message_handler(func=verificar)
def responder(mensagem):
    texto = """
    Olá eu sou o Jorge Tadeu, seu bot de erros:
    Escolha um sistema(Clique no item):
     /Chat_Bot
     /SAT_Central
     /Meu_INSS
     /SAG_Internet
     /Novo_SEC
Responder qualquer outra coisa não vai funcionar, clique em uma das opções"""
    bot.reply_to(mensagem, texto)
    
if __name__ == '__main__':  
    while True: 
        time.sleep(5)
        navegador.get("https://meu.inss.gov.br/#/login")
        time.sleep(10)
        i = 0
        for x in loc_elem:
            indice = loc_elem.index(x)
            teste_elem(x, indice)
            time.sleep(5)
            print(x)
            if len(erros):
                func_chat_bot()
                chatbot()
                func_email()
                func_screen_chatbot()
                i += 1
                break
            else:
                continue
        navegador.get("https://requerimento.inss.gov.br/")
        #pymsgbox.alert('Prossiga com o login', 'Alerta')
        time.sleep(3600)
        erros.clear()
        
pysmsgbox.alert('Terminou o código', 'Jorge Tadeu')
print(list_err_chatbot)
#SAG
bot.polling()


# In[ ]:




