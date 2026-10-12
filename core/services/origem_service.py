"""Identificacao conservadora de projetos de origem para avisos editoriais.

Nunca deduz a origem apenas pelo ano ou pelo fato de o PDF estar vazio.
"""
import re

_PADRAO_G27 = re.compile(r"(?<![a-z0-9])g27(?![a-z0-9])", re.IGNORECASE)


def origem_g27(imagem):
    """Detecta somente referencias explicitas a G27 nos metadados da imagem."""
    projeto = getattr(imagem, "projeto", None)
    projeto_nome = getattr(projeto, "nome", "") if projeto is not None else ""
    originais = getattr(imagem, "dados_fotoweb_originais", None)
    if not isinstance(originais, dict):
        originais = {}
    valores = (
        getattr(imagem, "retranca", ""),
        projeto_nome,
        originais.get("__projeto_editorial_pdf", ""),
    )
    return any(_PADRAO_G27.search(str(valor or "")) for valor in valores)
