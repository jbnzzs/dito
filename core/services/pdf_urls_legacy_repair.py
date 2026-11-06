"""Reparo estrito de URLs antigas de PDF no SharePoint do Dito!.

Somente o segmento do nome do PDF gNN_e2_ANOpNNNletra_mp.pdf e alterado.
Não faz chamadas ao SharePoint e não toca em historicos, status ou metadados.
"""
import re
from urllib.parse import unquote, urlsplit

from core.models import Imagem

_SITE = "ensinolivre.sharepoint.com"
# Exemplos reconhecidos:
#   /MAT/g28_e2_7p013m_mp.pdf
#   %2FESP%2Fg28%5Fe2%5F9p013e%5Fmp%2Epdf
_PDF_ANTIGO = re.compile(
    r"(?P<fronteira>/|%2f)"
    r"(?P<projeto>g\d+)"
    r"(?P<separador1>_|%5f)e2(?P<separador2>_|%5f)"
    r"(?=\d{1,2}p\d{3,4}[a-z]{1,3}(?:_|%5f)mp(?:\.|%2e)pdf(?:$|[?&#]|%26))",
    flags=re.IGNORECASE,
)


def corrigir_url_pdf_antiga(url):
    """Devolve URL corrigida ou a original se nao for inequivocamente elegivel."""
    if not url or not isinstance(url, str):
        return url
    try:
        endereco = urlsplit(url)
    except ValueError:
        return url
    if endereco.scheme != "https" or (endereco.hostname or "").lower() != _SITE:
        return url
    decodificado = unquote(url).lower()
    if "/sites/g25digital/" not in decodificado or "/pdfs/" not in decodificado:
        return url

    def trocar(correspondencia):
        return (
            correspondencia.group("fronteira")
            + correspondencia.group("projeto")
            + correspondencia.group("separador1")
            + "e002"
            + correspondencia.group("separador2")
        )

    corrigida, quantidade = _PDF_ANTIGO.subn(trocar, url)
    return corrigida if quantidade == 1 else url


def listar_candidatos():
    """Leitura do banco ativo deste ambiente, inclusive quando rodando no Render."""
    registros = Imagem.objects.filter(ativo=True, url_pdf__icontains="e2").values_list(
        "pk", "retranca", "url_pdf"
    ).order_by("pk")
    encontrados = []
    for pk, retranca, antiga in registros.iterator(chunk_size=500):
        nova = corrigir_url_pdf_antiga(antiga)
        if nova != antiga:
            encontrados.append((pk, retranca, antiga, nova))
    return encontrados
