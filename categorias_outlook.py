"""Categorias das Areas do Second Brain no Outlook."""

import json
import unicodedata
from urllib.parse import quote

import requests


GRAPH_BASE_URL = "https://graph.microsoft.com/v1.0"
REQUEST_TIMEOUT = 60

CATEGORIAS_AREA = {
    "malkuth": {
        "display_name": "🌍 Malkuth · Cotidiano & Vida Prática",
        "color": "preset16",
    },
    "yesod": {
        "display_name": "☽ Yesod · Cultura, Imaginação & Percepção",
        "color": "preset9",
    },
    "hod": {
        "display_name": "☿ Hod · Comunicação & Aprendizado",
        "color": "preset1",
    },
    "netzach": {
        "display_name": "♀ Netzach · Relacionamentos & Vínculos",
        "color": "preset4",
    },
    "tiphereth": {
        "display_name": (
            "☉ Tiphereth · Eu, Propósito & Desenvolvimento Interior"
        ),
        "color": "preset3",
    },
    "geburah": {
        "display_name": "♂ Geburah · Saúde, Corpo & Disciplina",
        "color": "preset0",
    },
    "chesed": {
        "display_name": "♃ Chesed · Finanças & Patrimônio",
        "color": "preset7",
    },
    "binah": {
        "display_name": "♄ Binah · Carreira & Estrutura Profissional",
        "color": "preset15",
    },
}


class CategoriaOutlookError(RuntimeError):
    def __init__(
        self,
        status_code,
        method,
        url,
        response_text,
    ):
        self.status_code = status_code
        self.method = method
        self.url = url
        self.response_text = response_text

        mensagem = (
            "Microsoft Graph retornou erro ao processar "
            "uma categoria do Outlook.\n"
            f"HTTP: {status_code}\n"
            f"Metodo: {method}\n"
            f"URL: {url}\n"
            f"Resposta: {response_text}"
        )

        super().__init__(mensagem)


def normalizar_texto(valor):
    texto = unicodedata.normalize(
        "NFKD",
        str(valor or ""),
    )

    texto = "".join(
        caractere
        for caractere in texto
        if not unicodedata.combining(caractere)
    )

    return " ".join(
        texto.casefold().strip().split()
    )


def chave_da_area(area):
    texto = normalizar_texto(area)

    for chave in CATEGORIAS_AREA:
        if chave in texto:
            return chave

    return None


def categoria_da_area(area):
    chave = chave_da_area(area)

    if not chave:
        return None

    return CATEGORIAS_AREA[chave]["display_name"]


def categorias_da_area(area):
    categoria = categoria_da_area(area)

    if not categoria:
        return []

    return [categoria]


def headers_graph(access_token):
    return {
        "Authorization": f"Bearer {access_token}",
        "Accept": "application/json",
        "Content-Type": "application/json",
    }


def conteudo_resposta(response):
    try:
        return json.dumps(
            response.json(),
            indent=2,
            ensure_ascii=False,
        )
    except ValueError:
        return response.text or "(resposta vazia)"


def validar_resposta(response):
    if response.ok:
        return

    raise CategoriaOutlookError(
        status_code=response.status_code,
        method=response.request.method,
        url=response.url,
        response_text=conteudo_resposta(response),
    )


def requisicao_graph(
    method,
    url,
    access_token,
    json_body=None,
    timeout=REQUEST_TIMEOUT,
):
    response = requests.request(
        method=method,
        url=url,
        headers=headers_graph(access_token),
        json=json_body,
        timeout=timeout,
    )

    validar_resposta(response)

    if response.status_code == 204:
        return None

    if not response.content:
        return None

    return response.json()


def listar_categorias(
    access_token,
    timeout=REQUEST_TIMEOUT,
):
    url = (
        f"{GRAPH_BASE_URL}/me/outlook/masterCategories"
        "?$top=100"
    )

    categorias = []

    while url:
        resultado = requisicao_graph(
            method="GET",
            url=url,
            access_token=access_token,
            timeout=timeout,
        )

        if not resultado:
            break

        categorias.extend(
            resultado.get("value", [])
        )

        url = resultado.get("@odata.nextLink")

    por_nome_exato = {}
    por_nome_normalizado = {}

    for categoria in categorias:
        display_name = categoria.get("displayName")

        if not display_name:
            continue

        por_nome_exato[display_name] = categoria

        por_nome_normalizado[
            normalizar_texto(display_name)
        ] = categoria

    return {
        "lista": categorias,
        "por_nome_exato": por_nome_exato,
        "por_nome_normalizado": por_nome_normalizado,
    }


def localizar_categoria(
    categorias,
    display_name,
):
    categoria = categorias[
        "por_nome_exato"
    ].get(display_name)

    if categoria:
        return categoria

    return categorias[
        "por_nome_normalizado"
    ].get(
        normalizar_texto(display_name)
    )


def criar_categoria(
    access_token,
    display_name,
    color,
    timeout=REQUEST_TIMEOUT,
):
    return requisicao_graph(
        method="POST",
        url=(
            f"{GRAPH_BASE_URL}"
            "/me/outlook/masterCategories"
        ),
        access_token=access_token,
        json_body={
            "displayName": display_name,
            "color": color,
        },
        timeout=timeout,
    )


def atualizar_cor_categoria(
    access_token,
    categoria_id,
    color,
    timeout=REQUEST_TIMEOUT,
):
    categoria_id_codificado = quote(
        str(categoria_id),
        safe="",
    )

    return requisicao_graph(
        method="PATCH",
        url=(
            f"{GRAPH_BASE_URL}"
            "/me/outlook/masterCategories/"
            f"{categoria_id_codificado}"
        ),
        access_token=access_token,
        json_body={
            "color": color,
        },
        timeout=timeout,
    )


def criar_categoria_com_recuperacao(
    access_token,
    display_name,
    color,
    timeout,
):
    try:
        return criar_categoria(
            access_token=access_token,
            display_name=display_name,
            color=color,
            timeout=timeout,
        )

    except CategoriaOutlookError as error:
        if error.status_code != 400:
            raise

        print(
            "O Graph rejeitou a criacao. "
            "Verificando se a categoria ja existe..."
        )

        categorias_atualizadas = listar_categorias(
            access_token=access_token,
            timeout=timeout,
        )

        categoria_existente = localizar_categoria(
            categorias_atualizadas,
            display_name,
        )

        if categoria_existente:
            print(
                "Categoria localizada depois "
                "da nova consulta: "
                f"{display_name}"
            )

            return categoria_existente

        raise


def garantir_categorias(
    access_token,
    timeout=REQUEST_TIMEOUT,
):
    print()
    print("=" * 70)
    print("CATEGORIAS DAS AREAS NO OUTLOOK")
    print("=" * 70)

    categorias = listar_categorias(
        access_token=access_token,
        timeout=timeout,
    )

    print(
        "Categorias encontradas antes "
        f"da sincronizacao: {len(categorias['lista'])}"
    )

    criadas = 0
    atualizadas = 0
    preservadas = 0
    avisos = 0

    for configuracao in CATEGORIAS_AREA.values():
        nome = configuracao["display_name"]
        cor = configuracao["color"]

        print()
        print(f"Processando categoria: {nome}")
        print(f"Cor desejada: {cor}")

        atual = localizar_categoria(
            categorias,
            nome,
        )

        if not atual:
            atual = criar_categoria_com_recuperacao(
                access_token=access_token,
                display_name=nome,
                color=cor,
                timeout=timeout,
            )

            if atual:
                criadas += 1

                print(
                    f"Categoria criada ou recuperada: {nome}"
                )
            else:
                avisos += 1

                print(
                    "AVISO: o Graph nao retornou "
                    f"a categoria criada: {nome}"
                )

            categorias = listar_categorias(
                access_token=access_token,
                timeout=timeout,
            )

            continue

        cor_atual = atual.get("color")

        if cor_atual == cor:
            preservadas += 1

            print(
                f"Categoria preservada: {nome}"
            )

            continue

        categoria_id = atual.get("id")

        if not categoria_id:
            avisos += 1

            print(
                "AVISO: categoria encontrada "
                "sem ID. Cor nao atualizada: "
                f"{nome}"
            )

            continue

        atualizar_cor_categoria(
            access_token=access_token,
            categoria_id=categoria_id,
            color=cor,
            timeout=timeout,
        )

        atualizadas += 1

        print(
            "Cor da categoria atualizada: "
            f"{nome} | {cor_atual} -> {cor}"
        )

    print()
    print("=" * 70)
    print("RESUMO DAS CATEGORIAS")
    print("=" * 70)
    print(f"Categorias criadas/recuperadas: {criadas}")
    print(f"Categorias atualizadas: {atualizadas}")
    print(f"Categorias preservadas: {preservadas}")
    print(f"Avisos: {avisos}")
