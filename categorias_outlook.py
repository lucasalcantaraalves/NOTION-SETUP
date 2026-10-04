"""
Categorias visuais das Areas do Second Brain no Outlook.

Este modulo:
- identifica a Area de um registro;
- retorna a categoria correspondente;
- cria as categorias ausentes no Outlook;
- preserva as categorias que ja estiverem corretas;
- atualiza a cor quando necessario.
"""

from __future__ import annotations

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


def normalizar_texto(valor: object) -> str:
    """
    Normaliza o texto para facilitar a identificacao da Area.

    Remove acentos, converte para minusculas e normaliza espacos.
    """

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


def chave_da_area(area: object) -> str | None:
    """
    Identifica a chave interna correspondente a Area recebida.

    Exemplos:
    ♀ Netzach
    Netzach
    ♀ Netzach · Relacionamentos & Vinculos
    """

    texto = normalizar_texto(area)

    for chave in CATEGORIAS_AREA:
        if chave in texto:
            return chave

    return None


def categoria_da_area(area: object) -> str | None:
    """
    Retorna o nome da categoria correspondente a Area.
    """

    chave = chave_da_area(area)

    if not chave:
        return None

    return CATEGORIAS_AREA[chave]["display_name"]


def categorias_da_area(area: object) -> list"""
    Retorna a lista de categorias que deve ser aplicada
    ao evento do Outlook.
    """

    categoria = categoria_da_area(area)

    if not categoria:
        return []

    return [categoria]


def headers_graph(access_token: str) -> dict[str, str]:
    """
    Monta os cabecalhos utilizados nas chamadas ao Microsoft Graph.
    """

    return {
        "Authorization": f"Bearer {access_token}",
        "Accept": "application/json",
        "Content-Type": "application/json",
    }


def listar_categorias(
    access_token: str,
    timeout: int = REQUEST_TIMEOUT,
) -> dict[str, dict]:
    """
    Lista as categorias existentes no Outlook.
    """

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
    access_token: str,
    display_name: str,
    color: str,
    timeout: int = REQUEST_TIMEOUT,
) -> dict:
    """
    Cria uma categoria no Outlook.
    """

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
    access_token: str,
    categoria_id: str,
    color: str,
    timeout: int = REQUEST_TIMEOUT,
) -> dict:
    """
    Atualiza a cor de uma categoria existente.
    """

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
    access_token: str,
    timeout: int = REQUEST_TIMEOUT,
) -> None:
    """
    Garante que as categorias das oito Areas existam no Outlook.

    Categorias ausentes sao criadas.
    Categorias existentes com a cor correta sao preservadas.
    Categorias existentes com outra cor sao atualizadas.
    """

    print()
    print("=" * 70)
    print("CATEGORIAS DAS AREAS NO OUTLOOK")
    print("=" * 70)

    categorias_existentes = listar_categorias(
        access_token=access_token,
        timeout=timeout,
    )

    criadas = 0
    atualizadas = 0
    preservadas = 0

    for configuracao in CATEGORIAS_AREA.values():
        display_name = configuracao["display_name"]
        color = configuracao["color"]

        categoria_existente = categorias_existentes.get(
            display_name
        )

        if not categoria_existente:
            criar_categoria(
                access_token=access_token,
                display_name=display_name,
                color=color,
                timeout=timeout,
            )

            criadas += 1

            print(
                f"Categoria criada: {display_name}"
            )

            continue

        cor_atual = categoria_existente.get(
            "color"
        )

        if cor_atual == color:
            preservadas += 1

            print(
                f"Categoria preservada: {display_name}"
            )

            continue

        categoria_id = categoria_existente.get(
            "id"
        )

        if not categoria_id:
            print(
                "Categoria encontrada sem ID. "
                "Nao foi possivel atualizar a cor: "
                f"{display_name}"
            )

            continue

        atualizar_cor_categoria(
            access_token=access_token,
            categoria_id=categoria_id,
            color=color,
            timeout=timeout,
        )

        atualizadas += 1

        print(
            f"Cor da categoria atualizada: {display_name}"
        )

    print()
    print(f"Categorias criadas: {criadas}")
    print(f"Categorias atualizadas: {atualizadas}")
    print(f"Categorias preservadas: {preservadas}")
