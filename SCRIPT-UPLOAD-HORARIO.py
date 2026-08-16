import os
from google.oauth2 import service_account
from googleapiclient.discovery import build
from googleapiclient.http import MediaFileUpload
from datetime import datetime

# IDs configurados corretamente
SHEET_ID = "1UN08EyAA6gj8tiUXLWpzoRNiucaHrr_xHtu8j6UnLfw"
PASTA_DESTINO_ID = "13-RfMhN9tNJgrya8YDyIBMnBGQq73fvS"
FILE_NAME = "log_execucao.csv"

# Autenticação usando o arquivo gerado pelo GitHub Actions
SCOPES = [
    'https://www.googleapis.com/auth/spreadsheets',
    'https://www.googleapis.com/auth/drive'
]

creds = service_account.Credentials.from_service_account_file(
    os.environ.get('GOOGLE_APPLICATION_CREDENTIALS', 'credentials.json'), 
    scopes=SCOPES
)

service_drive = build('drive', 'v3', credentials=creds)
service_sheets = build('sheets', 'v4', credentials=creds)

def criar_log():
    timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    with open(FILE_NAME, "w", encoding="utf-8") as f:
        f.write(f"Timestamp,{timestamp}\n")

    print("Arquivo CSV criado com sucesso!")
    return FILE_NAME

def upload_e_atualizar(file_path, service_drive, service_sheets):
    # Adicionamos o 'parents' com o ID da pasta nos metadados
    file_metadata = {
        'name': FILE_NAME,
        'parents': [PASTA_DESTINO_ID] 
    }
    
    media = MediaFileUpload(file_path, mimetype='text/csv')
    file = service_drive.files().create(body=file_metadata, media_body=media, fields='id').execute()

    print("Upload para o Drive realizado com ID:", file.get('id'))
    
    # Importar para o Sheets (append na primeira aba)
    with open(file_path, "r", encoding="utf-8") as f:
        valores = [line.strip().split(',') for line in f.readlines()]
    
    service_sheets.spreadsheets().values().append(
        spreadsheetId=SHEET_ID,
        range="Página1!A1",
        valueInputOption="USER_ENTERED",
        body={"values": valores}
    ).execute()

    print("Planilha atualizada com sucesso!")


# Bloco principal que executa as funções em ordem
if __name__ == "__main__":
    # 1. Cria o arquivo CSV localmente
    arquivo_gerado = criar_log()
    
    # 2. Faz o upload para o Drive e atualiza a planilha do Sheets
    upload_e_atualizar(arquivo_gerado, service_drive, service_sheets)
    
    print("Processo concluído com sucesso!")dentials=creds)
