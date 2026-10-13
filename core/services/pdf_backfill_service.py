"""Preenche somente links SharePoint ausentes; nao altera fluxo editorial.

A fonte dos metadados e o relatorio FotoWeb enviado pelo coordenador.
Cada execucao e limitada a 1000 linhas para evitar timeout no Render.
"""
from collections import Counter

from django.db import transaction
from django.db.models import Q

from core.models import Imagem
from .excel_service import _valor_excel
from .sharepoint_service import _url_pdf_dados_fotoweb

MAXIMO_LINHAS_PDFS = 1000


def preparar_pdfs_existentes(projeto, linhas, editorial, scriba):
    """Somente SELECT; lista candidatos para previsualizacao e aplicacao."""
    motivos = Counter()
    por_retranca = {}
    ambiguas = set()

    for linha in linhas:
        retranca = str(_valor_excel(linha, "retranca") or "").strip()
        if not retranca:
            motivos["retranca_vazia"] += 1
            continue

        url, motivo = _url_pdf_dados_fotoweb(editorial, scriba, linha, retranca)
        if not url:
            motivos[motivo or "sem_link"] += 1
            continue

        if retranca in por_retranca and por_retranca[retranca] != url:
            ambiguas.add(retranca)
        por_retranca[retranca] = url

    for retranca in ambiguas:
        por_retranca.pop(retranca, None)
    motivos["retranca_ambigua"] += len(ambiguas)

    candidatas = []
    encontradas = 0
    preenchidas = 0
    if por_retranca:
        imagens = (
            Imagem.objects.filter(
                projeto=projeto,
                ativo=True,
                retranca__in=por_retranca,
            )
            .only("id", "retranca", "url_pdf")
            .order_by("pk")
        )
        for imagem in imagens.iterator(chunk_size=250):
            encontradas += 1
            if imagem.url_pdf:
                preenchidas += 1
                continue
            candidatas.append({
                "id": imagem.pk,
                "retranca": imagem.retranca,
                "url": por_retranca[imagem.retranca],
            })

    return {
        "candidatas": candidatas,
        "encontradas": encontradas,
        "ja_preenchidas": preenchidas,
        "sem_correspondencia": max(0, len(por_retranca) - encontradas),
        "motivos": dict(motivos),
        "amostras": candidatas[:12],
    }


def aplicar_pdfs_existentes(projeto, candidatas):
    """UPDATE estritamente de url_pdf vazio, revalidado sob bloqueio de linha."""
    urls_por_pk = {item["id"]: item["url"] for item in candidatas}
    ids = sorted(urls_por_pk)
    atualizadas = 0

    for inicio in range(0, len(ids), 200):
        lote_ids = ids[inicio:inicio + 200]
        with transaction.atomic():
            imagens = list(
                Imagem.objects.select_for_update()
                .filter(projeto=projeto, ativo=True, pk__in=lote_ids)
                .filter(Q(url_pdf="") | Q(url_pdf__isnull=True))
                .only("id", "url_pdf")
            )
            for imagem in imagens:
                imagem.url_pdf = urls_por_pk[imagem.pk]
            if imagens:
                Imagem.objects.bulk_update(imagens, ["url_pdf"], batch_size=100)
            atualizadas += len(imagens)

    return atualizadas
