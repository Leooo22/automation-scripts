import pyautogui as pg
import time
import os
import dotenv as load_dotenv 
# Carregar variáveis de ambiente
load_dotenv()

Past_host = os.getenv('Location_Past_Xlsx')

def copiar_dados_do_sheets():
    try:
        pg.hotkey('ctrl', 'a')
        time.sleep(1)
        pg.hotkey('ctrl', 'c')
        time.sleep(2)
        print("Dados copiados com sucesso!")
    except Exception as e:
        print(f"Erro ao copiar os dados do Sheets: {e}")
        exit()

def colar_no_excel():
    try:
        pg.press('win')
        time.sleep(1)
        pg.write('Past_host')
        time.sleep(1)
        pg.press('enter')
        time.sleep(5)
        pg.hotkey('ctrl', 'v')
        time.sleep(2)
        print("Dados colados no Excel com sucesso!")
    except Exception as e:
        print(f"Erro ao colar os dados no Excel: {e}")
        exit()
