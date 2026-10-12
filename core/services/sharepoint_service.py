"""Construção de links de PDFs SharePoint. Sem acesso ao banco de dados."""

import re
import unicodedata
from urllib.parse import quote

from .excel_service import _valor_excel

# ============================================================
# PDF — SHAREPOINT
# ============================================================

SHAREPOINT_SITE_URL = "https://ensinolivre.sharepoint.com/sites/G25Digital"
SHAREPOINT_BIBLIOTECA = "Shared Documents"
SHAREPOINT_PASTA_PDFS = "PDFs"
PDF_TIPO_MATERIAL = "mp"

PDF_COMPONENTES_SHAREPOINT = {
    "ARTE": "ART",
    "ART": "ART",
    "HISTORIA": "HIS",
    "HIS": "HIS",
}

# DITO_SHAREPOINT_V4 - aliases conhecidos de pastas no SharePoint
PDF_COMPONENTES_SHAREPOINT.update({
    'MATEMATICA': 'MAT',
    'MAT': 'MAT',
    'CIENCIA': 'CIE',
    'CIENCIAS': 'CIE',
    'CIENCIASDANATUREZA': 'CIE',
    'CIE': 'CIE',
    'GEOGRAFIA': 'GEO',
    'GEO': 'GEO',
})


def _normalizar_codigo_sharepoint(valor):
    texto = str(valor or "").strip()
    texto = unicodedata.normalize("NFKD", texto).encode("ascii", "ignore").decode("ascii")
    return re.sub(r"[^A-Za-z0-9]+", "", texto).upper()


def _codigo_componente_pdf_sharepoint(valor):
    normalizado = _normalizar_codigo_sharepoint(valor)
    if not normalizado:
        return ""
    return PDF_COMPONENTES_SHAREPOINT.get(normalizado, "")


def _montar_url_pdf_sharepoint(retranca, componente):
    retranca_normalizada = str(retranca or "").strip().lower()
    if not retranca_normalizada:
        return ""

    dados_pagina = re.search(
        r"(?<!\d)(?P<ano>\d{1,2})p(?P<projeto_scriba>\d{3})(?P<letra>[a-z])",
        retranca_normalizada,
    )
    dados_editoriais = re.search(
        r"(?P<projeto>g\d+)[_-](?P<editora>e\d+)",
        retranca_normalizada,
    )
    componente_codigo = _codigo_componente_pdf_sharepoint(componente)

    if not dados_pagina or not dados_editoriais or not componente_codigo:
        return ""

    projeto = dados_editoriais.group("projeto")
    editora = dados_editoriais.group("editora")
    ano = dados_pagina.group("ano")
    projeto_scriba = dados_pagina.group("projeto_scriba")
    letra = dados_pagina.group("letra")

    nome_pdf = (
        f"{projeto}_{editora}_{ano}p{projeto_scriba}{letra}_{PDF_TIPO_MATERIAL}.pdf"
    )

    segmentos = [
        SHAREPOINT_BIBLIOTECA,
        projeto,
        SHAREPOINT_PASTA_PDFS,
        componente_codigo,
        nome_pdf,
    ]
    caminho = "/".join(quote(segmento, safe="") for segmento in segmentos)
    return f"{SHAREPOINT_SITE_URL}/{caminho}?web=1"

# DITO_SHAREPOINT_V4 - metadados do Excel como fonte primaria para o PDF.
# DITO_SHAREPOINT_METADADOS: URL calculada com dados do Excel + formulário.

# DITO_SHAREPOINT_PASTAS_CONFIRMADAS_20261008: pastas, códigos de PDF e anos confirmados no SharePoint
PDF_COMPONENTES_SHAREPOINT.update({'PORTUGUES': 'POR', 'POR': 'POR', 'ESPANHOL': 'ESP', 'ESP': 'ESP', 'INGLES': 'ING', 'ING': 'ING', 'EDUCACAODIGITAL': 'EDIG', 'EDDIGITAL': 'EDIG', 'EDIG': 'EDIG', 'EDUCACAOFISICA': 'EFIS', 'EDFISICA': 'EFIS', 'EFIS': 'EFIS', 'PDT': 'PDT', 'ARTE': 'ART', 'ART': 'ART', 'CIENCIA': 'CIE', 'CIENCIAS': 'CIE', 'CIENCIASDANATUREZA': 'CIE', 'CIE': 'CIE', 'GEOGRAFIA': 'GEO', 'GEO': 'GEO', 'HISTORIA': 'HIS', 'HIS': 'HIS', 'MATEMATICA': 'MAT', 'MAT': 'MAT'})
DITO_SP_SUFIXO_ARQUIVO = {'ART': 'a', 'CIE': 'c', 'EDIG': 'ed', 'EFIS': 'ef', 'ESP': 'e', 'GEO': 'g', 'HIS': 'h', 'ING': 'i', 'MAT': 'm', 'PDT': 't', 'POR': 'p'}
DITO_SP_ANOS_CONFIRMADOS = {'EDIG': (6, 7, 8, 9), 'EFIS': (6,), 'ESP': (6, 7, 8, 9), 'ING': (6, 7, 8, 9), 'PDT': (6, 7, 8, 9), 'POR': (6, 7, 8, 9)}

def _montar_url_pdf_metadados_dito(editorial, scriba, sigla, ano):
    """Gera URL somente para combinações reconhecidas de pasta e ano."""
    from urllib.parse import quote as _quote_dito

    partes = str(editorial or "").strip().upper().split("_")
    if (len(partes) < 2
            or not re.fullmatch(r"G\d+", partes[0])
            or not re.fullmatch(r"E\d+", partes[1])):
        return ""

    projeto_scriba = str(scriba or "").strip().lower()
    if not re.fullmatch(r"p?\d{3,4}", projeto_scriba):
        return ""
    numero = "p" + projeto_scriba.removeprefix("p").zfill(3)

    pasta = str(sigla or "").strip().upper()
    sufixo = DITO_SP_SUFIXO_ARQUIVO.get(pasta)
    if not sufixo:
        return ""

    serie = str(ano or "").strip()
    if not re.fullmatch(r"\d{1,2}", serie):
        return ""
    ano_num = int(serie)
    if not 1 <= ano_num <= 12:
        return ""
    anos_confirmados = DITO_SP_ANOS_CONFIRMADOS.get(pasta)
    if anos_confirmados is not None and ano_num not in anos_confirmados:
        return ""

    prefixo = (partes[0] + "_" + partes[1]).lower()
    nome_pdf = f"{prefixo}_{ano_num}{numero}{sufixo}_mp.pdf"
    caminho = "/".join(
        _quote_dito(str(segmento), safe="")
        for segmento in (
            SHAREPOINT_BIBLIOTECA,
            partes[0].lower(),
            SHAREPOINT_PASTA_PDFS,
            pasta,
            nome_pdf,
        )
    )
    return f"{SHAREPOINT_SITE_URL.rstrip('/')}/{caminho}?web=1"


def _url_pdf_dados_fotoweb(editorial, scriba, dados, retranca=""):
    componente = _valor_excel(dados, "disciplina", "componente", "componente_curricular")
    volume = _valor_excel(dados, "volume", "volume_ano_modulo")
    codigo = PDF_COMPONENTES_SHAREPOINT.get(_normalizar_codigo_sharepoint(componente), "")
    if not codigo:
        return "", "componente_nao_mapeado"
    # Não inferir série de números longos, como 2029 -> 20.
    anos = re.findall(r"(?<!\d)\d{1,2}(?!\d)", str(volume or ""))
    if len(anos) != 1 or not (1 <= int(anos[0]) <= 12):
        return "", "ano_ausente_ou_ambiguo"
    # Retranca é apenas verificação adicional de conflito inequívoco.
    m = re.search(r"(?i)(?:^|[_-])p\d{3,4}[_-](cie|mat|his|art|geo)(?=$|[_-])", str(retranca or ""))
    if m and m.group(1).upper() != codigo:
        return "", "conflito_excel_retranca"
    url = _montar_url_pdf_metadados_dito(editorial, scriba, codigo, anos[0])
    return url, ("ok" if url else "url_nao_gerada")
