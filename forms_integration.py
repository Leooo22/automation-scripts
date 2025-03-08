import pyautogui as pg
import time
import os
from dotenv import load_dotenv
import pyperclip
from excel_interaction import copiar_dados_do_sheets

# Carregar variáveis de ambiente
load_dotenv()

email = os.getenv('Login_usuario')
password = os.getenv('User_Password')
url = os.getenv('Url_Sheets')

def abrir_navegador():
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

def fazer_login():
    try:
        pg.hotkey('ctrl', 'l')  # Selecionar barra de endereço
        pg.write('https://accounts.google.com/')
        pg.press('enter')
        time.sleep(5)

        if not email or not password:
            raise ValueError("Credenciais não carregadas corretamente.")

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

def acessar_drive(url):
    try:
        if not url:
            raise ValueError("URL da planilha não carregada corretamente.")
        
        pg.hotkey('ctrl', 't')  # Nova guia
        time.sleep(2)
        pg.write(url)
        pg.press('enter')
        time.sleep(5)
        print("Planilha aberta com sucesso!")
    except Exception as e:
        print(f"Erro ao acessar a planilha: {e}")

def copiar_links_coluna_c():
    try:
        links = []
        pg.hotkey('ctrl', 'home')  # Ir para o topo da planilha
        time.sleep(1)

        # Move para a coluna C
        pg.press('right', presses=2, interval=0.1)
        time.sleep(1)

        while True:
            pg.hotkey('ctrl', 'c')  # Copia o conteúdo da célula
            time.sleep(1)
            link = pyperclip.paste()  # Pega o conteúdo copiado

            if not link.startswith("http"):  # Para se não houver mais links
                break

            links.append(link)
            pg.press('down')  # Vai para a próxima célula
            time.sleep(1)

        print("Links copiados da coluna C:", links)
        return links
    except Exception as e:
        print(f"Erro ao copiar os links da coluna C: {e}")
        return []

def clicar_nos_links_com_botao_direito(links):
    for link in links:
        try:
            pg.hotkey('ctrl', 't')  # Nova guia
            time.sleep(2)
            pyperclip.copy(link)
            pg.hotkey('ctrl', 'v')  # Cola o link
            pg.press('enter')
            time.sleep(5)  # Espera carregar

            print(f"Acessando o link: {link}")

            # Copiar dados da planilha aberta
            copiar_dados_do_sheets()  

            # Fecha a aba depois de copiar os dados
            pg.hotkey('ctrl', 'w')
            time.sleep(1)
            pg.hotkey('ctrl', 'shift', 'tab')  # Volta para a planilha original
        except Exception as e:
            print(f"Erro ao acessar o link: {e}")

# Executar o script
if __name__ == "__main__":
    abrir_navegador()
    fazer_login()
    acessar_drive()
    links = copiar_links_coluna_c()
    clicar_nos_links_com_botao_direito(links)
