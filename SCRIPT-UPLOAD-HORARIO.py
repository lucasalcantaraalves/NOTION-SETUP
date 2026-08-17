import os
import json
import requests
import pandas as pd
from datetime import datetime
from google.oauth2 import service_account
from googleapiclient.discovery import build

# Configurações e validações de segurança das variáveis de ambiente
SHEET_ID = "1UN08EyAA6gj8tiUXLWpzoRNiucaHrr_xHtu8j6UnLfw"

NOTION_TOKEN = os.environ.get('NOTION_TOKEN')
if not NOTION_TOKEN:
    raise ValueError("A variável de ambiente NOTION_TOKEN não foi encontrada!")

creds_json_str = os.environ.get('GOOGLE_CREDENTIALS_JSON')
if not creds_json_str:
    raise ValueError("A variável de ambiente GOOGLE_APPLICATION_CREDENTIALS_JSON não foi encontrada! Verifique os Secrets no GitHub.")

creds_json = json.loads(creds_json_str)
creds = service_account.Credentials.from_service_account_info(creds_json, scopes=['https://www.googleapis.com/auth/spreadsheets'])
service_sheets = build('sheets', 'v4', credentials=creds)

# === CLASSE NOTION SYNC (Responsável por buscar e formatar os dados do Notion) ===
class NotionSync:
    def __init__(self):
        self.headers = {
            "Authorization": f"Bearer {NOTION_TOKEN}",
            "Content-Type": "application/json",
            "Notion-Version": "2022-06-28"
        }

    def notion_total_view(self, database_id):
        url = f"https://api.notion.com/v1/databases/{database_id}/query"
        pages = []
        has_more = True
        start_cursor = None

        while has_more:
            body = {}
            if start_cursor:
                body["start_cursor"] = start_cursor
            
            response = requests.post(url, headers=self.headers, json=body)
            if response.status_code != 200:
                raise Exception(f"Erro ao buscar dados do Notion: {response.text}")
            
            data = response.json()
            pages.extend(data.get("results", []))
            has_more = data.get("has_more", False)
            start_cursor = data.get("next_cursor")
            
        return pages

    def format_dataframe(self, results):
        parsed_data = []
        for page in results:
            props = page.get("properties", {})
            row = {}
            for prop_name, prop_val in props.items():
                ptype = prop_val.get("type")
                val = None
                
                if ptype == "title":
                    texts = prop_val.get("title", [])
                    val = "".join([t.get("plain_text", "") for t in texts]) if texts else ""
                elif ptype == "rich_text":
                    texts = prop_val.get("rich_text", [])
                    val = "".join([t.get("plain_text", "") for t in texts]) if texts else ""
                elif ptype == "select":
                    sel = prop_val.get("select")
                    val = sel.get("name") if sel else ""
                elif ptype == "multi_select":
                    msel = prop_val.get("multi_select", [])
                    val = ", ".join([s.get("name", "") for s in msel]) if msel else ""
                elif ptype == "status":
                    status = prop_val.get("status")
                    val = status.get("name") if status else ""
                elif ptype == "date":
                    date_obj = prop_val.get("date")
                    val = date_obj.get("start") if date_obj else ""
                elif ptype in ["number", "checkbox"]:
                    val = prop_val.get(ptype)
                elif ptype == "url":
                    val = prop_val.get("url")
                elif ptype == "email":
                    val = prop_val.get("email")
                elif ptype == "phone_number":
                    val = prop_val.get("phone_number")
                else:
                    val = str(prop_val)
                
                row[prop_name] = val
            parsed_data.append(row)
            
        return pd.DataFrame(parsed_data)

# === FUNÇÃO DE SINCRONIZAÇÃO ===
def sincronizar_db(db_id, aba_nome):
    print(f"🔄 Iniciando sincronização para a aba: {aba_nome}")
    print("🔄 Buscando dados do Notion...")
    
    nsync = NotionSync()
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

# === EXECUÇÃO DAS DUAS BASES ===
# Base 1 (Master / Página1)
sincronizar_db('312b40ec7cd4807fa77dc62a474bc6b4', 'Master')

# Base 2 (Embarques / Página 2)
sincronizar_db('30ab40ec7cd48071aca9fc8d5ac81e6a', 'Embarques')
