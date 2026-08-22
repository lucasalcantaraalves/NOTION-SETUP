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

SCOPES = ['https://www.googleapis.com/auth/calendar']
creds = service_account.Credentials.from_service_account_info(creds_json, scopes=SCOPES)
service_calendar = build('calendar', 'v3', credentials=creds)

DATABASE_ID = "312b40ec7cd4807fa77dc62a474bc6b4"
CALENDAR_ID = "e14lucasdejesus@gmail.com" # Substitua pelo seu ID de calendário correto

headers = {
    "Authorization": f"Bearer {NOTION_TOKEN}",
    "Content-Type": "application/json",
    "Notion-Version": "2022-06-28"
}

def sincronizar_com_calendar():
    print("🔄 Sincronizando itens do Notion com o Google Calendar...")
    
    # Busca todas as páginas da base que possuem Data preenchida
    url = f"https://api.notion.com/v1/databases/{DATABASE_ID}/query"
    payload = {
        "filter": {
            "property": "Due Date",
            "date": {
                "is_not_empty": True
            }
        }
    }
    
    response = requests.post(url, headers=headers, json=payload)
    if response.status_code != 200:
        raise Exception(f"Erro ao buscar dados do Notion: {response.text}")
        
    pages = response.json().get("results", [])
    print(f"✅ {len(pages)} itens com data encontrados para verificar/sincronizar.")
    
    for page in pages:
        page_id = page["id"]
        props = page.get("properties", {})
        
        # 1. Extração do Título
        nome_prop = props.get("Name", {}).get("title", [])
        titulo_evento = "".join([t.get("plain_text", "") for t in nome_prop]) if nome_prop else "Evento sem nome"
        
        # 2. Extração da Data
        data_prop = props.get("Due Date", {}).get("date")
        if not data_prop or not data_prop.get("start"):
            continue
            
        start_str = data_prop.get("start")
        start_dt = datetime.fromisoformat(start_str.replace('Z', '+00:00'))
        end_dt = start_dt + timedelta(hours=1)

        # 3. Extração do Event ID existente no Notion
        event_id_prop = props.get("Calendar Event ID", {}).get("rich_text", [])
        calendar_event_id = "".join([t.get("plain_text", "") for t in event_id_prop]) if event_id_prop else ""

        # Montar o corpo do evento
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
            if not calendar_event_id:
                # CRIAR NOVO EVENTO
                created_event = service_calendar.events().insert(calendarId=CALENDAR_ID, body=evento_body).execute()
                new_event_id = created_event.get('id')
                print(f"📅 Evento criado: {titulo_evento}")
                
                # Salvar o ID do Google Calendar e marcar o checkbox no Notion
                update_url = f"https://api.notion.com/v1/pages/{page_id}"
                update_payload = {
                    "properties": {
                        "Calendar Event ID": {
                            "rich_text": [{"text": {"content": new_event_id}}]
                        },
                        "Status Calendar": {
                            "checkbox": True
                        }
                    }
                }
                requests.patch(update_url, headers=headers, json=update_payload)
                
            else:
                # ATUALIZAR EVENTO EXISTENTE (Caso tenha mudado a data ou nome no Notion)
                service_calendar.events().update(
                    calendarId=CALENDAR_ID, 
                    eventId=calendar_event_id, 
                    body=evento_body
                ).execute()
                print(f"🔄 Evento atualizado: {titulo_evento}")
                
        except Exception as e:
            print(f"❌ Erro ao processar o evento '{titulo_evento}': {str(e)}")

if __name__ == "__main__":
    sincronizar_com_calendar()
