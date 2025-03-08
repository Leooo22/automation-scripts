import os
from forms_integration import abrir_navegador, fazer_login, acessar_drive
from excel_interaction import copiar_dados_do_sheets, colar_no_excel
from dotenv import load_dotenv

# Carregar variáveis de ambiente
load_dotenv()
url = os.getenv('Url_Sheets')

if not url:
    print("Erro: URL da planilha não encontrada no arquivo .env")
    exit()
else:
    print(f"URL da planilha carregada: {url}")

def main():
    abrir_navegador()
    fazer_login()
    acessar_drive(url)
    copiar_dados_do_sheets()
    colar_no_excel()

if __name__ == "__main__":
    main()
