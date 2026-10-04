"""
Categorias visuais das Áreas do Second Brain no Outlook.

O nome da categoria identifica a Área.
A cor da categoria fornece a identidade visual no calendário.
"""

from __future__ import annotations

import unicodedata
from urllib.parse import quote

import requests


GRAPH_BASE_URL = "https://graph.microsoft.com/v1.0"


CATEGORIAS_AREA = {
    "malkuth": {
        "displayName": "🌍 Malkuth · Cotidiano & Vida Prática",
        "color": "preset16",
    },
    "yesod": {
        "displayName": "☽ Yesod · Cultura, Imaginação & Percepção",
        "color": "preset9",
    },
    "hod": {
        "displayName": "☿ Hod · Comunicação & Aprendizado",
        "color": "preset1",
    },
    "netzach": {
        "displayName": "♀ Netzach · Relacionamentos & Vínculos",
        "color": "preset4",
    },
    "tiphereth": {
        "displayName": "☉ Tiphereth · Eu, Propósito & Desenvolvimento",
        "color": "preset3",
    },
    "geburah": {
        "displayName": "♂ Geburah · Saúde, Corpo & Disciplina",
        "color": "preset0",
    },
    "chesed": {
        "displayName": "♃ Chesed · Finanças & Patrimônio",
        "color": "preset7",
    },
    "binah": {
        "displayName": "♄ Binah · Carreira & Estrutura Profissional",
        "color": "preset15",
    },
}


def normalizar_texto(valor: object) -> str:
    """
    Remove acentos, converte para minúsculas e normaliza espaços.
    """

    texto = unicodedata.normalize("NFKD", str(valor or ""))

    texto = "".join(
        caractere
        for caractere in texto
        if not unicodedata.combining(caractere)
    )

    return " ".join(texto.lower().strip().split())


def chave_da_area(area: object) -> str | None:
    """
    Identifica a chave da Área a partir do texto recebido do Excel.

    Exemplos aceitos:
    - ♀ Netzach
    - Netzach
    - ♀ Netzach · Relacionamentos & Vínculos
    """

    texto = normalizar_texto(area)

    for chave in CATEGORIAS_AREA:
        if chave in texto:
            return chave

    return None


def categorias_da_area(area: object) -> list"""
    Retorna a categoria que deve ser aplicada ao evento do Outlook.
    """

    chave = chave_da_area(area)

    if not chave:
        return []

    return [CATEGORIAS_AREA[chave]["displayName"]]


def headers_graph(access_token: str) -> dict[str, str]:
    """
    Monta os cabeçalhos utilizados nas chamadas ao Microsoft Graph.
    """

    return {
        "Authorization": f"Bearer {access_token}",
        "Content-Type": "application/json",
    }


def listar_categorias(
    access_token: str,
    timeout: int = 30,
) -> dict[str, dict]:
    """
    Obtém as categorias existentes na conta do Outlook.
    """

    response = requests.get(
        f"{GRAPH_BASE_URL}/me/outlook/masterCategories",
        headers=headers_graph(access_token),
        timeout=timeout,
    )

    response.raise_for_status()

    categorias = response.json().get("value", [])

    return {
        categoria.get("displayName", ""): categoria
        for categoria in categorias
        if categoria.get("displayName")
    }


def criar_categoria(
    access_token: str,
    display_name: str,
    color: str,
    timeout: int = 30,
) -> None:
    """
    Cria uma nova categoria no Outlook.
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


def atualizar_cor_categoria(
    access_token: str,
    category_id: str,
    color: str,
    timeout: int = 30,
) -> None:
    """
    Atualiza a cor de uma categoria existente.
    """

    category_id_codificado = quote(str(category_id), safe="")

    response = requests.patch(
        (
            f"{GRAPH_BASE_URL}/me/outlook/"
            f"masterCategories/{category_id_codificado}"
        ),
        headers=headers_graph(access_token),
        json={
            "color": color,
        },
        timeout=timeout,
    )

    response.raise_for_status()


def garantir_categorias(
    access_token: str,
    timeout: int = 30,
) -> None:
    """
    Garante que todas as categorias das Áreas existam no Outlook.

    Se a categoria já existir com outra cor, a cor será atualizada.
    """

    categorias_existentes = listar_categorias(
        access_token=access_token,
        timeout=timeout,
    )

    for configuracao in CATEGORIAS_AREA.values():
        display_name = configuracao["displayName"]
        color = configuracao["color"]

        categoria_existente = categorias_existentes.get(display_name)

        if not categoria_existente:
            criar_categoria(
                access_token=access_token,
                display_name=display_name,
                color=color,
                timeout=timeout,
            )

            print(f"Categoria criada no Outlook: {display_name}")
            continue

        cor_atual = categoria_existente.get("color")

        if cor_atual == color:
            continue

        category_id = categoria_existente.get("id")

        if not category_id:
            print(
                "Categoria encontrada sem ID. "
                f"Não foi possível atualizar a cor: {display_name}"
            )
            continue

        atualizar_cor_categoria(
            access_token=access_token,
            category_id=category_id,
            color=color,
            timeout=timeout,
        )

        print(f"Cor da categoria atualizada: {display_name}")
