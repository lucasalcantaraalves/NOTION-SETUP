
"""Categorias das Areas do Second Brain no Outlook."""

import re
import unicodedata
from html import unescape
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
        "display_name": "☽ Yesod · Cultura | Imaginação & Percepção",
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
        "display_name": "☉ Tiphereth · Eu | Propósito & Desenvolvimento Interior",
        "color": "preset3",
    },
    "geburah": {
        "display_name": "♂ Geburah · Saúde | Corpo & Disciplina",
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
    texto = unicodedata.normalize("NFKD", str(valor or ""))
    texto = "".join(
        caractere
        for caractere in texto
        if not unicodedata.combining(caractere)
    )
    return " ".join(texto.lower().strip().split())


def chave_da_area(area):
    texto = normalizar_texto(area)

    for chave in CATEGORIAS_AREA:
        if re.search(rf"\b{re.escape(chave)}\b", texto):
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


def listar_categorias(access_token, timeout=REQUEST_TIMEOUT):
    response = requests.get(
        f"{GRAPH_BASE_URL}/me/outlook/masterCategories",
        headers=headers_graph(access_token),
        timeout=timeout,
    )
    response.raise_for_status()

    categorias = response.json().get("value", [])

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
    categoria_id_codificado = quote(str(categoria_id), safe="")

    response = requests.patch(
        (
            f"{GRAPH_BASE_URL}/me/outlook/"
            f"masterCategories/{categoria_id_codificado}"
        ),
        headers=headers_graph(access_token),
        json={"color": color},
        timeout=timeout,
    )
    response.raise_for_status()

    if not response.content:
        return {}

    return response.json()


def garantir_categorias(access_token, timeout=REQUEST_TIMEOUT):
    """Cria as categorias ausentes e corrige as cores configuradas."""
    print()
    print("=" * 70)
    print("CATEGORIAS DAS AREAS NO OUTLOOK")
    print("=" * 70)

    existentes = listar_categorias(access_token, timeout)

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
            print(f"Categoria criada: {nome}")
            continue

        if atual.get("color") == cor:
            preservadas += 1
            print(f"Categoria preservada: {nome}")
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
        print(f"Cor da categoria atualizada: {nome}")

    print()
    print(f"Categorias criadas: {criadas}")
    print(f"Categorias atualizadas: {atualizadas}")
    print(f"Categorias preservadas: {preservadas}")


def texto_simples(html):
    """Converte a descrição HTML em texto, preservando limites de linha."""
    texto = str(html or "")

    texto = re.sub(
        r"(?i)<\s*br\s*/?\s*>|</\s*(?:p|div|li|tr)\s*>",
        "\n",
        texto,
    )
    texto = re.sub(r"<[^>]+>", " ", texto)
    texto = unescape(texto)

    return texto


def identificar_area_na_descricao(descricao):
    """
    Identifica somente uma área declarada explicitamente.

    Exemplos aceitos:
      Área: Netzach
      Area: ♀ Netzach · Relacionamentos & Vínculos
      Área - Geburah
    """
    texto = texto_simples(descricao)

    for linha in texto.splitlines():
        correspondencia = re.match(
            r"^\s*(?:área|area)\s*[:\-–]\s*(.*?)\s*$",
            linha,
            flags=re.IGNORECASE,
        )

        if not correspondencia:
            continue

        area_declarada = correspondencia.group(1)
        chave = chave_da_area(area_declarada)

        if chave:
            return chave

    return None


def listar_eventos(access_token, timeout=REQUEST_TIMEOUT):
    """Lista eventos do calendário padrão, seguindo a paginação do Graph."""
    url = (
        f"{GRAPH_BASE_URL}/me/events"
        "?$select=id,subject,body,categories"
        "&$top=100"
    )

    headers = headers_graph(access_token)
    eventos = []

    while url:
        response = requests.get(
            url,
            headers=headers,
            timeout=timeout,
        )
        response.raise_for_status()

        dados = response.json()
        eventos.extend(dados.get("value", []))
        url = dados.get("@odata.nextLink")

    return eventos


def atualizar_categorias_evento(
    access_token,
    evento_id,
    categorias,
    timeout=REQUEST_TIMEOUT,
):
    """Atualiza as categorias do evento sem modificar os demais campos."""
    evento_id_codificado = quote(str(evento_id), safe="")

    response = requests.patch(
        f"{GRAPH_BASE_URL}/me/events/{evento_id_codificado}",
        headers=headers_graph(access_token),
        json={"categories": categorias},
        timeout=timeout,
    )
    response.raise_for_status()

    if not response.content:
        return {}

    return response.json()


def sincronizar_areas_eventos(
    access_token,
    timeout=REQUEST_TIMEOUT,
    modo_simulacao=True,
):
    """
    Aplica a categoria correspondente à área declarada na descrição.

    modo_simulacao=True:
        Apenas mostra as alterações propostas.

    modo_simulacao=False:
        Aplica as alterações no Outlook.

    Eventos sem uma área explícita válida permanecem inalterados.
    Categorias que não pertencem às oito áreas são preservadas.
    """
    print()
    print("=" * 70)
    print("SINCRONIZACAO AUTOMATICA DAS AREAS DOS EVENTOS")
    print("=" * 70)

    existentes = listar_categorias(access_token, timeout)
    eventos = listar_eventos(access_token, timeout)

    nomes_areas = {
        config["display_name"]
        for config in CATEGORIAS_AREA.values()
    }

    atualizados = 0
    ja_corretos = 0
    sem_area = 0
    categoria_ausente = 0
    simulados = 0

    for evento in eventos:
        assunto = evento.get("subject") or "(sem assunto)"
        evento_id = evento.get("id")

        body = evento.get("body") or {}
        descricao = body.get("content", "")

        area = identificar_area_na_descricao(descricao)

        if not area:
            sem_area += 1
            print(f"[SEM AREA] {assunto}")
            continue

        nome_categoria = CATEGORIAS_AREA[area]["display_name"]

        if nome_categoria not in existentes:
            categoria_ausente += 1
            print(
                f"[CATEGORIA AUSENTE] {assunto}: "
                f"{nome_categoria}"
            )
            continue

        categorias_atuais = evento.get("categories") or []

        # Retira as categorias das áreas que estejam incorretas,
        # preserva todas as categorias que não pertencem ao Second Brain
        # e aplica apenas a categoria da área declarada.
        outras_categorias = [
            categoria
            for categoria in categorias_atuais
            if categoria not in nomes_areas
        ]

        categorias_novas = outras_categorias + [nome_categoria]

        if set(categorias_novas) == set(categorias_atuais):
            ja_corretos += 1
            print(f"[OK] {assunto} -> {nome_categoria}")
            continue

        if not evento_id:
            print(f"[ERRO: SEM ID] {assunto}")
            continue

        if modo_simulacao:
            simulados += 1
            print(
                f"[SIMULACAO] {assunto}\n"
                f"    Atual: {categorias_atuais}\n"
                f"    Nova:  {categorias_novas}"
            )
            continue

        try:
            atualizar_categorias_evento(
                access_token,
                evento_id,
                categorias_novas,
                timeout,
            )
        except requests.RequestException as erro:
            print(f"[ERRO] {assunto}: {erro}")
            continue

        atualizados += 1
        print(f"[ATUALIZADO] {assunto} -> {nome_categoria}")

    print()
    print(f"Eventos examinados: {len(eventos)}")
    print(f"Eventos já corretos: {ja_corretos}")
    print(f"Eventos sem área explícita: {sem_area}")
    print(f"Categorias ausentes: {categoria_ausente}")
    print(f"Alterações simuladas: {simulados}")
    print(f"Eventos atualizados: {atualizados}")
