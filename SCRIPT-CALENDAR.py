import os
import json
import requests
from datetime import datetime, timedelta
from google.oauth2 import service_account
from googleapiclient.discovery import build

# Configurações de Ambiente
NOTION_TOKEN = os.environ.get('NOTION_TOKEN')
if not NOTION_TOKEN:
    raise ValueError("A variável de ambiente NOTION_TOKEN não foi encontrada!")

creds_json_str = os.environ.get('GOOGLE_CREDENTIALS_JSON')
if not creds_json_str:
    raise ValueError("A variável de ambiente GOOGLE_CREDENTIALS_JSON não foi encontrada!")

creds_json = json.loads(creds_json_str)

# ATENÇÃO: É necessário incluir o escopo do Calendar além do Sheets
SCOPES = ['https://www.googleapis.com/auth/calendar']
creds = service_account.Credentials.from_service_account_info(creds_json, scopes=SCOPES)
service_calendar = build('calendar', 'v3', credentials=creds)

# ID da database do Notion que você quer monitorar
DATABASE_ID = "312b40ec7cd4807fa77dc62a474bc6b4" # Substitua se for outra base
CALENDAR_ID = "primary" # 'primary' é o calendário principal da conta da service account ou o ID do seu Google Calendar

headers = {
    "Authorization": f"Bearer {NOTION_TOKEN}",
    "Content-Type": "application/json",
    "Notion-Version": "2022-06-28"
}

def sincronizar_com_calendar():
    print("🔄 Buscando itens pendentes no Notion...")
    
    # Query para buscar apenas onde o checkbox está desmarcado (False)
    url = f"https://api.notion.com/v1/databases/{DATABASE_ID}/query"
    payload = {
        "filter": {
            "property": "Status Calendar",
            "checkbox": {
                "equals": False
            }
        }
    }
    
    response = requests.post(url, headers=headers, json=payload)
    if response.status_code != 200:
        raise Exception(f"Erro ao buscar dados do Notion: {response.text}")
        
    pages = response.json().get("results", [])
    print(f"✅ {len(pages)} novos itens encontrados para sincronizar com o Calendar.")
    
    for page in pages:
        page_id = page["id"]
        props = page.get("properties", {})
        
        # 1. Extração do Título
        nome_prop = props.get("Name", {}).get("title", [])
        titulo_evento = "".join([t.get("plain_text", "") for t in nome_prop]) if nome_prop else "Evento sem nome"
        
        # 2. Extração da Data
        data_prop = props.get("Due Date", {}).get("date")
        if not data_prop or not data_prop.get("start"):
            print(f"⚠️ Item '{titulo_evento}' ignorado: Sem data definida.")
            continue
            
        start_str = data_prop.get("start")
        
        # Converte a string do Notion para objeto datetime (assume 00:00:00 se não houver hora)
        start_dt = datetime.fromisoformat(start_str.replace('Z', '+00:00'))
        
        # Define o fim como 1 hora após o início
        end_dt = start_dt + timedelta(hours=1)

        # 3. Montar o corpo do evento com fuso horário e duração de 1 hora
        evento_body = {
            'summary': titulo_evento,
            'start': {
                'dateTime': start_dt.isoformat(),
                'timeZone': 'America/Sao_Paulo'
            },
            'end': {
                'dateTime': end_dt.isoformat(),
                'timeZone': 'America/Sao_Paulo'
            }
        }
        
        try:
            service_calendar.events().insert(calendarId=CALENDAR_ID, body=evento_body).execute()
            print(f"📅 Evento criado no Calendar: {titulo_evento} ({start_str})")
            
            # 4. Marcar o checkbox como True no Notion para não processar de novo
            update_url = f"https://api.notion.com/v1/pages/{page_id}"
            update_payload = {
                "properties": {
                    "Status Calendar": {
                        "checkbox": True
                    }
                }
            }
            patch_resp = requests.patch(update_url, headers=headers, json=update_payload)
            if patch_resp.status_code == 200:
                print(f"✔️ Checkbox marcado como True no Notion para: {titulo_evento}")
            else:
                print(f"❌ Erro ao atualizar checkbox no Notion: {patch_resp.text}")
                
        except Exception as e:
            print(f"❌ Erro ao processar o evento '{titulo_evento}': {str(e)}")

if __name__ == "__main__":
    sincronizar_com_calendar()
