import pyautogui as pg
import time
import os
import openpyxl
from dotenv import load_dotenv

# Carregar variáveis de ambiente
load_dotenv()

# Recuperar credenciais de forma segura
email = os.getenv('Login_usuario')
password = os.getenv('User_Password')

# Verificar se as credenciais foram carregadas corretamente
if not email or not password:
    print("Erro: Credenciais não encontradas no arquivo .env")
    exit()

def abrir_navegador():
    # Abre o navegador Chrome em modo anônimo
    try:
        pg.press('win')
        time.sleep(1)
        pg.write('chrome')
        time.sleep(1)
        pg.press('enter')
        time.sleep(3)
        pg.hotkey('ctrl', 'shift', 'n')  # Abrir guia anônima
        time.sleep(2)
        print("Navegador aberto com sucesso!")
    except Exception as e:
        print(f"Erro ao abrir o navegador: {e}")
        exit()

def fazer_login():
    # Faz login no Google
    try:
        # Acessar a página de login do Google
        pg.hotkey('ctrl', 'l')  # Selecionar barra de endereço
        pg.write('https://accounts.google.com/')
        pg.press('enter')
        time.sleep(5)

        # Digitar e-mail e senha
        pg.write(email)
        time.sleep(1)
        pg.press('enter')
        time.sleep(3)
        pg.write(password)
        time.sleep(2)
        pg.press('enter')
        time.sleep(5)

        print("Login realizado com sucesso!")
    except Exception as e:
        print(f"Erro ao fazer login: {e}")
        exit()

def acessar_drive():
    # Abre o Google Drive e entra na planilha 
    try:
        # Acessar Google Drive
        pg.hotkey('ctrl', 't')  # Nova guia
        time.sleep(2)
        pg.write('https://drive.google.com/drive/home')
        pg.press('enter')
        time.sleep(5)
        print("Google Drive aberto com sucesso!")

        # Acessar planilha do Google Sheets
        pg.hotkey('ctrl', 't')  # Nova guia
        time.sleep(2)
        pg.write('https://docs.google.com/spreadsheets/d/1ODhNDCKoPAtlPRWPHtyHTbQ4UZJH74mpAqZBoI2LKTM/edit#gid=311446733')
        pg.press('enter')
        time.sleep(5)
        print("Planilha aberta com sucesso!")

    except Exception as e:
        print(f"Erro ao acessar o Google Drive e a planilha: {e}")
        exit()

def copiar_dados_do_sheets():
    # Seleciona e copia os dados da planilha
    try:
        # Selecionar toda a planilha (Ctrl + A)
        pg.hotkey('ctrl', 'a')
        time.sleep(1)

        # Copiar os dados (Ctrl + C)
        pg.hotkey('ctrl', 'c')
        time.sleep(2)
        print("Dados copiados com sucesso!")

    except Exception as e:
        print(f"Erro ao copiar os dados do Sheets: {e}")
        exit()

def colar_no_excel():
    # Abre o Excel e cola os dados copiados
    try:
        # Abrir o Excel (caso esteja no menu iniciar)
        pg.press('win')
        time.sleep(1)
        pg.write('Excel')
        time.sleep(1)
        pg.press('enter')
        time.sleep(5)

        # Criar nova planilha (Ctrl + N)
        pg.hotkey('ctrl', 'n')
        time.sleep(2)

        # Colar os dados (Ctrl + V)
        pg.hotkey('ctrl', 'v')
        time.sleep(2)
        print("Dados colados no Excel com sucesso!")

    except Exception as e:
        print(f"Erro ao colar os dados no Excel: {e}")
        exit()

def main():
    # Fluxo de execução do Script
    abrir_navegador()
    fazer_login()
    acessar_drive()
    copiar_dados_do_sheets()
    colar_no_excel()

if __name__ == "__main__":
    main()
