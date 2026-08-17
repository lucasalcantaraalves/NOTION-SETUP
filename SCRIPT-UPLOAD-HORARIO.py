import os
import requests
import pandas as pd
from datetime import datetime
from google.oauth2 import service_account
from googleapiclient.discovery import build

# === CONFIGURAÇÕES ===
SHEET_ID = "1UN08EyAA6gj8tiUXLWpzoRNiucaHrr_xHtu8j6UnLfw"
DATABASE_ID = '312b40ec7cd4807fa77dc62a474bc6b4'

# Puxa o token com segurança do ambiente do GitHub Actions
NOTION_TOKEN = os.environ.get('NOTION_TOKEN')
if not NOTION_TOKEN:
    raise ValueError("A variável de ambiente NOTION_TOKEN não foi configurada!")

# Autenticação Google Sheets
SCOPES = [
    'https://www.googleapis.com/auth/spreadsheets'
]

creds = service_account.Credentials.from_service_account_file(
    os.environ.get('GOOGLE_APPLICATION_CREDENTIALS', 'credentials.json'), 
    scopes=SCOPES
)
service_sheets = build('sheets', 'v4', credentials=creds)


class NotionSync:
    def __init__(self):
        self.headers = {
            "Authorization": "Bearer " + NOTION_TOKEN,
            "Notion-Version": "2022-06-28",
            "Content-Type": "application/json"
        }

    def notion_db_view(self, database_id, cursor=None):
        payload = {}
        if cursor:
            payload['start_cursor'] = cursor
        url = f"https://api.notion.com/v1/databases/{database_id}/query"
        response = requests.post(url, headers=self.headers, json=payload)
        if response.status_code != 200:
            raise Exception(f"Erro {response.status_code}: {response.text}")
        return response.json()

    def notion_total_view(self, database_id):
        cursor, pages = None, []
        while True:
            response = self.notion_db_view(database_id, cursor)
            pages += response['results']
            cursor = response.get('next_cursor', None)
            if not cursor:
                break
        return pages

    def format_dataframe(self, results):
        data = []
        for item in results:
            properties = item.get("properties", {})
            row = {"Page ID": item["id"]}
            for key, value in properties.items():
                tipo = value.get("type")
                try:
                    if tipo == "relation":
                        row[key] = ", ".join([rel.get("id") for rel in value.get("relation", [])])
                    elif tipo == "people":
                        row[key] = ", ".join([p.get("id", "") for p in value.get("people", [])])
                    elif tipo == "rich_text":
                        row[key] = ''.join([t.get("text", {}).get("content", "") for t in value.get("rich_text", [])])
                    elif tipo == "title":
                        row[key] = ''.join([t.get("text", {}).get("content", "") for t in value.get("title", [])])
                    elif tipo == "checkbox":
                        row[key] = str(value.get("checkbox", False))
                    elif tipo == "select":
                        row[key] = value["select"].get("name", "") if value.get("select") else ""
                    elif tipo == "multi_select":
                        row[key] = ", ".join([s.get("name", "") for s in value.get("multi_select", [])])
                    elif tipo == "date":
                        row[key] = value.get("date", {}).get("start", "")
                    elif tipo == "number":
                        row[key] = value.get("number", 0)
                    elif tipo == "url":
                        row[key] = value.get("url", "")
                    elif tipo == "email":
                        row[key] = value.get("email", "")
                    elif tipo == "phone_number":
                        row[key] = value.get("phone_number", "")
                    elif tipo == "created_time":
                        row[key] = value.get("created_time", "")
                    elif tipo == "last_edited_time":
                        row[key] = value.get("last_edited_time", "")
                    elif tipo == "formula":
                        formula = value.get("formula", {})
                        row[key] = formula.get("string") or formula.get("number") or formula.get("boolean") or ""
                    else:
                        row[key] = str(value.get(tipo, ""))
                except:
                    row[key] = ""
            data.append(row)
        return pd.DataFrame(data)


def sincronizar_tudo():
    print("🔄 Buscando dados do Notion...")
    nsync = NotionSync()
    results = nsync.notion_total_view(DATABASE_ID)
    print(f"✅ {len(results)} páginas encontradas no Notion.")

    df_notion = nsync.format_dataframe(results)

    # Prepara os dados para inserir na planilha
    timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    
    # Monta a estrutura de linhas que irão para o Sheets
    valores_para_planilha = []
    
    # Linha 1: O Timestamp da execução
    valores_para_planilha.append([f"Ultima Atualizacao: {timestamp}"])
    valores_para_planilha.append([]) # Linha em branco para separar
    
    # Se houver dados no Notion, adiciona o cabeçalho e as linhas da tabela
    if not df_notion.empty:
        # Cabeçalho das colunas do Notion
        valores_para_planilha.append(list(df_notion.columns))
        # Linhas de dados do Notion convertidas para lista
        for _, row in df_notion.iterrows():
            valores_para_planilha.append(list(row.values))

    # Limpa os dados anteriores da aba "Página1"
    service_sheets.spreadsheets().values().clear(
        spreadsheetId=SHEET_ID, 
        range="Página1!A:Z"
    ).execute()
    print("🧹 Dados anteriores da planilha apagados.")

    # Envia tudo para o Google Sheets
    service_sheets.spreadsheets().values().append(
        spreadsheetId=SHEET_ID,
        range="Página1!A1",
        valueInputOption="USER_ENTERED",
        body={"values": valores_para_planilha}
    ).execute()

    print("🚀 Planilha atualizada com sucesso com os dados do Notion!")


if __name__ == "__main__":
    sincronizar_tudo()
