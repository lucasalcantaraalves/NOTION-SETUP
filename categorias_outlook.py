"""Categorias das Areas do Second Brain no Outlook."""

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
        texto.lower().strip().split()
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


def listar_categorias(
    access_token,
    timeout=REQUEST_TIMEOUT,
):
    response = requests.get(
        f"{GRAPH_BASE_URL}/me/outlook/masterCategories",
        headers=headers_graph(access_token),
        timeout=timeout,
    )

    response.raise_for_status()

    categorias = response.json().get(
        "value",
        [],
    )

    return {
        categoria.get("displayName"): categoria
        for categoria in categorias
        if categoria.get("displayName")
    }


def criar_categoria(
    access_token,
    display_name,
    color,
    timeout=REQUEST_TIMEOUT,
):
    response = requests.post(
        f"{GRAPH_BASE_URL}/me/outlook/masterCategories",
        headers=headers_graph(access_token),
        json={
            "displayName": display_name,
            "color": color,
        },
        timeout=timeout,
    )

    response.raise_for_status()

    if not response.content:
        return {}

    return response.json()


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

    response = requests.patch(
        (
            f"{GRAPH_BASE_URL}/me/outlook/"
            f"masterCategories/{categoria_id_codificado}"
        ),
        headers=headers_graph(access_token),
        json={
            "color": color,
        },
        timeout=timeout,
    )

    response.raise_for_status()

    if not response.content:
        return {}

    return response.json()


def garantir_categorias(
    access_token,
    timeout=REQUEST_TIMEOUT,
):
    print()
    print("=" * 70)
    print("CATEGORIAS DAS AREAS NO OUTLOOK")
    print("=" * 70)

    existentes = listar_categorias(
        access_token,
        timeout,
    )

    criadas = 0
    atualizadas = 0
    preservadas = 0

    for configuracao in CATEGORIAS_AREA.values():
        nome = configuracao["display_name"]
        cor = configuracao["color"]
        atual = existentes.get(nome)

        if not atual:
            criar_categoria(
                access_token,
                nome,
                cor,
                timeout,
            )

            criadas += 1

            print(
                f"Categoria criada: {nome}"
            )

            continue

        if atual.get("color") == cor:
            preservadas += 1

            print(
                f"Categoria preservada: {nome}"
            )

            continue

        categoria_id = atual.get("id")

        if not categoria_id:
            print(
                "Categoria encontrada sem ID. "
                "Nao foi possivel atualizar a cor: "
                f"{nome}"
            )

            continue

        atualizar_cor_categoria(
            access_token,
            categoria_id,
            cor,
            timeout,
        )

        atualizadas += 1

        print(
            f"Cor da categoria atualizada: {nome}"
        )

    print()
    print(f"Categorias criadas: {criadas}")
    print(f"Categorias atualizadas: {atualizadas}")
    print(f"Categorias preservadas: {preservadas}")
