import base64
import hashlib
import json
import os
import sys
from pathlib import Path
from urllib.parse import quote

import msal
import requests
from cryptography.fernet import Fernet, InvalidToken

CLIENT_ID = os.environ.get("MS_CLIENT_ID")
CACHE_KEY_TEXT = os.environ.get("MS_CACHE_KEY")
if not CLIENT_ID:
    raise RuntimeError("Secret MS_CLIENT_ID nao encontrado.")
if not CACHE_KEY_TEXT:
    raise RuntimeError("Secret MS_CACHE_KEY nao encontrado.")

AUTHORITY = "https://login.microsoftonline.com/consumers"
SCOPES = ["User.Read", "Files.ReadWrite", "Calendars.ReadWrite"]
GRAPH = "https://graph.microsoft.com/v1.0"
TIMEOUT = 60
EXCEL_PATH = "/Microsoft Copilot Chat Files/Copilot Notebook Uploads/Second Brain - Dados.xlsx"
CACHE_PATH = Path(".auth/msal_cache.bin")
TABELA_DOCUMENTOS = "tbDocumentos"
TABELA_AGENDA = "tbAgenda"
STATUS_ATIVOS = {"ativo", "ativa"}
STATUS_INATIVOS = {"inativo", "inativa", "cancelado", "cancelada", "vencido", "vencida", "arquivado", "arquivada"}


class GraphError(Exception):
    def __init__(self, status, method, url, text=""):
        self.status = status
        self.method = method
        self.url = url
        self.text = text
        super().__init__(f"Graph HTTP {status}: {method} {url}")


def graph_request(method, endpoint, token, params=None, body=None, headers=None, aceitar_404=False):
    url = endpoint if endpoint.startswith("https://") else GRAPH + endpoint
    request_headers = {"Authorization": f"Bearer {token}", "Accept": "application/json"}
    if headers:
        request_headers.update(headers)
    response = requests.request(
        method, url, headers=request_headers, params=params, json=body, timeout=TIMEOUT
    )
    if aceitar_404 and response.status_code == 404:
        return None
    if not response.ok:
        raise GraphError(response.status_code, method, url, response.text)
    if response.status_code == 204 or not response.content:
        return None
    return response.json()


def chave_fernet(valor):
    raw = valor.encode("utf-8")
    try:
        Fernet(raw)
        return raw
    except Exception:
        return base64.urlsafe_b64encode(hashlib.sha256(raw).digest())


def carregar_cache():
    CACHE_PATH.parent.mkdir(parents=True, exist_ok=True)
    cache = msal.SerializableTokenCache()
    if not CACHE_PATH.exists():
        return cache
    try:
        texto = Fernet(chave_fernet(CACHE_KEY_TEXT)).decrypt(CACHE_PATH.read_bytes()).decode("utf-8")
        cache.deserialize(texto)
        print("Cache de autenticacao restaurado e descriptografado.")
        return cache
    except InvalidToken as erro:
        raise RuntimeError("Nao foi possivel descriptografar o cache MSAL.") from erro


def salvar_cache(cache):
    CACHE_PATH.parent.mkdir(parents=True, exist_ok=True)
    conteudo = Fernet(chave_fernet(CACHE_KEY_TEXT)).encrypt(cache.serialize().encode("utf-8"))
    CACHE_PATH.write_bytes(conteudo)
    print(f"Cache criptografado salvo em {CACHE_PATH}.")


def autenticar():
    cache = carregar_cache()
    app = msal.PublicClientApplication(CLIENT_ID, authority=AUTHORITY, token_cache=cache)
    resultado = None
    contas = app.get_accounts()
    if contas:
        print("Tentando autenticacao silenciosa...")
        resultado = app.acquire_token_silent(SCOPES, account=contas[0])
    if resultado and "access_token" in resultado:
        print("Autenticacao silenciosa concluida com sucesso.")
        salvar_cache(cache)
        return resultado["access_token"]

    print("Iniciando Device Code Flow...")
    fluxo = app.initiate_device_flow(scopes=SCOPES)
    if "user_code" not in fluxo:
        raise RuntimeError(json.dumps(fluxo, ensure_ascii=False))
    print(fluxo.get("message"))
    resultado = app.acquire_token_by_device_flow(fluxo)
    if "access_token" not in resultado:
        raise RuntimeError(json.dumps(resultado, ensure_ascii=False))
    salvar_cache(cache)
    return resultado["access_token"]


def localizar_excel(token):
    caminho = quote(EXCEL_PATH, safe="/")
    arquivo = graph_request(
        "GET", f"/me/drive/root:{caminho}", token,
        params={"$select": "id,name,lastModifiedDateTime"}
    )
    print(f"Excel encontrado: {arquivo['name']}")
    print(f"Item ID: {arquivo['id']}")
    return arquivo


def criar_sessao(token, item_id):
    resultado = graph_request(
        "POST", f"/me/drive/items/{item_id}/workbook/createSession", token,
        body={"persistChanges": True}
    )
    print("Sessao persistente do workbook criada.")
    return resultado["id"]


def fechar_sessao(token, item_id, session_id):
    graph_request(
        "POST", f"/me/drive/items/{item_id}/workbook/closeSession", token,
        headers={"workbook-session-id": session_id}
    )
    print("Sessao do workbook encerrada.")


def cabecalhos_tabela(token, item_id, session_id, tabela):
    resultado = graph_request(
        "GET", f"/me/drive/items/{item_id}/workbook/tables/{tabela}/headerRowRange", token,
        headers={"workbook-session-id": session_id}
    )
    valores = resultado.get("values", [])
    if not valores or not valores[0]:
        raise RuntimeError(f"A tabela {tabela} nao retornou cabecalhos.")
    return valores[0]


def linhas_tabela(token, item_id, session_id, tabela):
    endpoint = f"/me/drive/items/{item_id}/workbook/tables/{tabela}/rows"
    linhas = []
    params = {"$top": "200"}
    while endpoint:
        resultado = graph_request(
            "GET", endpoint, token, params=params,
            headers={"workbook-session-id": session_id}
        )
        params = None
        linhas.extend(resultado.get("value", []))
        endpoint = resultado.get("@odata.nextLink")
    return linhas


def registros_tabela(cabecalhos, linhas):
    registros = []
    for posicao, linha in enumerate(linhas):
        matriz = linha.get("values", [])
        valores = matriz[0] if matriz else []
        if not any(str(v or "").strip() for v in valores):
            continue
        registro = {
            cabecalho: valores[indice] if indice < len(valores) else ""
            for indice, cabecalho in enumerate(cabecalhos)
        }
        registros.append({
            "graph_index": linha.get("index", posicao),
            "valores": valores,
            "registro": registro,
        })
    return registros


def normalizar(valor):
    return str(valor or "").strip().casefold()


def preencher_linha(cabecalhos, dados):
    return [dados.get(cabecalho, "") for cabecalho in cabecalhos]


def atualizar_linha(token, item_id, session_id, tabela, graph_index, valores):
    graph_request(
        "PATCH",
        f"/me/drive/items/{item_id}/workbook/tables/{tabela}/rows/itemAt(index={graph_index})/range",
        token,
        body={"values": [valores]},
        headers={"workbook-session-id": session_id},
    )


def adicionar_linha(token, item_id, session_id, tabela, valores):
    resultado = graph_request(
        "POST", f"/me/drive/items/{item_id}/workbook/tables/{tabela}/rows/add", token,
        body={"index": None, "values": [valores]},
        headers={"workbook-session-id": session_id},
    )
    return resultado


def definir_campo(item, cabecalhos, campo, valor):
    if campo not in cabecalhos:
        raise RuntimeError(f"A coluna '{campo}' nao existe.")
    valores = list(item["valores"])
    while len(valores) < len(cabecalhos):
        valores.append("")
    valores[cabecalhos.index(campo)] = valor
    item["valores"] = valores
    item["registro"][campo] = valor


def id_agenda_padrao(documento_id):
    return f"DOC-{str(documento_id).strip()}-VENC"


def dados_agenda_do_documento(documento, agenda_id, status_agenda):
    documento_id = str(documento.get("ID") or "").strip()
    nome = str(documento.get("Nome") or documento_id).strip()
    vencimento = documento.get("Vencimento")
    referencia = str(documento.get("Documento") or "").strip()
    observacoes = str(documento.get("Observações") or "").strip()
    partes_observacao = [f"Gerado automaticamente a partir de tbDocumentos: {documento_id}."]
    if observacoes:
        partes_observacao.append(observacoes)
    return {
        "ID": agenda_id,
        "Nome": f"Vencimento · {nome}",
        "Área": documento.get("Área") or "",
        "Tipo": "Vencimento",
        "Data Início": vencimento or "",
        "Hora Início": "",
        "Data Fim": vencimento or "",
        "Hora Fim": "",
        "Dia Inteiro": "Sim",
        "Recorrência": "Nenhuma",
        "Status": status_agenda,
        "Lembrete": "Sim" if status_agenda == "Ativo" else "Não",
        "Política de Aviso": "180/90/30/7",
        "Referência": referencia or f"tbDocumentos:{documento_id}",
        "Observações": " ".join(partes_observacao),
        "Última atualização": documento.get("Última atualização") or "",
    }


def sincronizar_documentos(token, item_id, session_id):
    cab_docs = cabecalhos_tabela(token, item_id, session_id, TABELA_DOCUMENTOS)
    cab_agenda = cabecalhos_tabela(token, item_id, session_id, TABELA_AGENDA)
    docs = registros_tabela(cab_docs, linhas_tabela(token, item_id, session_id, TABELA_DOCUMENTOS))
    agenda = registros_tabela(cab_agenda, linhas_tabela(token, item_id, session_id, TABELA_AGENDA))

    agenda_por_id = {
        str(item["registro"].get("ID") or "").strip(): item
        for item in agenda
        if str(item["registro"].get("ID") or "").strip()
    }

    resumo = {"criados": 0, "atualizados": 0, "inativados": 0, "ignorados": 0, "erros": 0}
    print()
    print("=" * 70)
    print("SINCRONIZACAO tbDocumentos -> tbAgenda")
    print("=" * 70)
    print(f"Documentos encontrados: {len(docs)}")

    for item_doc in docs:
        doc = item_doc["registro"]
        doc_id = str(doc.get("ID") or "").strip()
        nome = str(doc.get("Nome") or "").strip()
        status_doc = normalizar(doc.get("Status"))
        vencimento = doc.get("Vencimento")
        agenda_id = str(doc.get("Registro de Agenda") or "").strip() or id_agenda_padrao(doc_id)

        print()
        print("-" * 70)
        print(f"Processando {doc_id} | {nome} | Status: {doc.get('Status')}")

        try:
            if not doc_id:
                raise RuntimeError("Documento sem ID.")

            existente = agenda_por_id.get(agenda_id)

            if status_doc in STATUS_ATIVOS and vencimento not in (None, ""):
                dados = dados_agenda_do_documento(doc, agenda_id, "Ativo")
                if existente:
                    outlook_id = existente["registro"].get("Outlook Event ID") or ""
                    dados["Outlook Event ID"] = outlook_id
                    nova_linha = preencher_linha(cab_agenda, dados)
                    atualizar_linha(
                        token, item_id, session_id, TABELA_AGENDA,
                        existente["graph_index"], nova_linha
                    )
                    existente["valores"] = nova_linha
                    existente["registro"].update(dados)
                    resumo["atualizados"] += 1
                    print(f"Registro de agenda atualizado: {agenda_id}")
                else:
                    dados["Outlook Event ID"] = ""
                    nova_linha = preencher_linha(cab_agenda, dados)
                    adicionar_linha(token, item_id, session_id, TABELA_AGENDA, nova_linha)
                    resumo["criados"] += 1
                    print(f"Registro de agenda criado: {agenda_id}")

                if str(doc.get("Registro de Agenda") or "").strip() != agenda_id:
                    definir_campo(item_doc, cab_docs, "Registro de Agenda", agenda_id)
                    atualizar_linha(
                        token, item_id, session_id, TABELA_DOCUMENTOS,
                        item_doc["graph_index"], item_doc["valores"]
                    )
                    print("Vinculo gravado em tbDocumentos.")

            elif existente and (status_doc in STATUS_INATIVOS or vencimento in (None, "")):
                dados = dados_agenda_do_documento(doc, agenda_id, "Inativo")
                dados["Outlook Event ID"] = existente["registro"].get("Outlook Event ID") or ""
                nova_linha = preencher_linha(cab_agenda, dados)
                atualizar_linha(
                    token, item_id, session_id, TABELA_AGENDA,
                    existente["graph_index"], nova_linha
                )
                resumo["inativados"] += 1
                print(f"Registro de agenda marcado como Inativo: {agenda_id}")

            else:
                resumo["ignorados"] += 1
                print("Documento sem vencimento monitoravel ou sem vinculo existente. Ignorado.")

        except Exception as erro:
            resumo["erros"] += 1
            print(f"ERRO em {doc_id or '[sem ID]'}: {erro}")

    print()
    print("=" * 70)
    print("RESUMO tbDocumentos -> tbAgenda")
    print("=" * 70)
    print(f"Registros de agenda criados: {resumo['criados']}")
    print(f"Registros de agenda atualizados: {resumo['atualizados']}")
    print(f"Registros de agenda inativados: {resumo['inativados']}")
    print(f"Documentos ignorados: {resumo['ignorados']}")
    print(f"Erros: {resumo['erros']}")
    if resumo["erros"]:
        raise RuntimeError("Uma ou mais linhas falharam. Veja os logs acima.")


def main():
    token = None
    item_id = None
    session_id = None
    try:
        token = autenticar()
        usuario = graph_request("GET", "/me", token, params={"$select": "mail,userPrincipalName"})
        print(f"Conta conectada: {usuario.get('mail') or usuario.get('userPrincipalName')}")
        arquivo = localizar_excel(token)
        item_id = arquivo["id"]
        session_id = criar_sessao(token, item_id)
        sincronizar_documentos(token, item_id, session_id)
        fechar_sessao(token, item_id, session_id)
        session_id = None
        print()
        print("SINCRONIZACAO DOCUMENTOS -> AGENDA CONCLUIDA COM SUCESSO.")
    except GraphError as erro:
        print(f"ERRO GRAPH {erro.status}: {erro.method} {erro.url}")
        print(erro.text)
        sys.exit(1)
    except Exception as erro:
        print(f"ERRO: {erro!r}")
        sys.exit(1)
    finally:
        if token and item_id and session_id:
            try:
                fechar_sessao(token, item_id, session_id)
            except Exception:
                pass


if __name__ == "__main__":
    main()
