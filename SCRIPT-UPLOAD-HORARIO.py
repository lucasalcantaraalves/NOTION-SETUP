import os
import json
import pandas as pd
from datetime import datetime
from google.oauth2 import service_account
from googleapiclient.discovery import build

# Configurações de Autenticação
SHEET_ID = "1UN08EyAA6gj8tiUXLWpzoRNiucaHrr_xHtu8j6UnLfw"
NOTION_TOKEN = os.environ.get('NOTION_TOKEN')
creds_json = json.loads(os.environ.get('GOOGLE_APPLICATION_CREDENTIALS_JSON'))
creds = service_account.Credentials.from_service_account_info(creds_json, scopes=['https://www.googleapis.com/auth/spreadsheets'])
service_sheets = build('sheets', 'v4', credentials=creds)

def sincronizar_db(db_id, aba_nome):
    print(f"🔄 Iniciando sincronização para a aba: {aba_nome}")
    print("🔄 Buscando dados do Notion...")
    
    nsync = NotionSync() # Mantendo sua classe original
    results = nsync.notion_total_view(db_id)
    print(f"✅ {len(results)} páginas encontradas no Notion.")

    df = nsync.format_dataframe(results).fillna("")
    
    # Prepara os dados
    timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    valores = [[f"Ultima Atualizacao: {timestamp}"]]
    valores.append([])
    if not df.empty:
        valores.append(list(df.columns))
        for _, row in df.iterrows():
            valores.append([str(val) for val in row.values])
            
    # Limpa dados anteriores
    service_sheets.spreadsheets().values().clear(spreadsheetId=SHEET_ID, range=f"{aba_nome}!A:Z").execute()
    print("🧹 Dados anteriores da planilha apagados.")

    # Envia para a planilha
    service_sheets.spreadsheets().values().append(
        spreadsheetId=SHEET_ID, range=f"{aba_nome}!A1", 
        valueInputOption="USER_ENTERED", body={"values": valores}
    ).execute()
    print("🚀 Planilha atualizada com sucesso com os dados do Notion!")

    # Salva o CSV
    nome_csv = f"log_{aba_nome}.csv"
    df.to_csv(nome_csv, index=False, encoding='utf-8-sig')
    print("📁 Arquivo CSV salvo com sucesso!")

# EXECUÇÃO DAS DUAS BASES
# Base 1 (Página 1)
sincronizar_db('312b40ec7cd4807fa77dc62a474bc6b4', 'Página1')

# Base 2 (Página 2)
sincronizar_db('30ab40ec7cd48071aca9fc8d5ac81e6a', 'Página2')
