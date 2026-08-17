import os
from google.oauth2 import service_account
from googleapiclient.discovery import build
from datetime import datetime

# IDs configurados
SHEET_ID = "1UN08EyAA6gj8tiUXLWpzoRNiucaHrr_xHtu8j6UnLfw"
FILE_NAME = "log_execucao.csv"

# Autenticação
SCOPES = [
    'https://www.googleapis.com/auth/spreadsheets'
]

creds = service_account.Credentials.from_service_account_file(
    os.environ.get('GOOGLE_APPLICATION_CREDENTIALS', 'credentials.json'), 
    scopes=SCOPES
)

service_sheets = build('sheets', 'v4', credentials=creds)

def criar_e_atualizar_planilha():
    timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    
    # 1. Cria o log localmente
    with open(FILE_NAME, "w", encoding="utf-8") as f:
        f.write(f"Timestamp,{timestamp}\n")

    print("Arquivo de log local criado com sucesso!")

    # 2. Lê os dados gerados
    with open(FILE_NAME, "r", encoding="utf-8") as f:
        valores = [line.strip().split(',') for line in f.readlines()]
    
    # 3. Envia direto para a planilha do Google Sheets
    service_sheets.spreadsheets().values().append(
        spreadsheetId=SHEET_ID,
        range="Página1!A1",
        valueInputOption="USER_ENTERED",
        body={"values": valores}
    ).execute()

    print("Planilha atualizada com sucesso!")

if __name__ == "__main__":
    criar_e_atualizar_planilha()
    print("Processo concluído com sucesso!")
