"""Exportação CSV em blocos para o relatório geral do Dito!.

Não altera as exportações XLSX nem os filtros existentes.
"""

import csv

from django.contrib.auth.decorators import login_required
from django.db.models import Q
from django.http import StreamingHttpResponse
from django.utils import timezone

from .relatorios_colunas import colunas_exportacao_dito, selecionar_valores_dito
from .relatorios_datas import imagens_exportacao_dito, datas_formatadas_dito


COLUNAS_GERAL = [
    "Retranca",
    "Projeto",
    "Obra",
    "Componente curricular",
    "Volume/Ano/Módulo",
    "Capítulo/Unidade",
    "Etapa",
    "Lote",
    "Data prevista do lote",
    "Data efetiva do lote",
    "Status",
    "Responsável atual",
    "Descritor",
    "Revisor",
    "Coordenador",
    "Pagamento descritor",
    "Pagamento revisor",
    "Prazo da imagem",
    "Cadastrada em",
    "Atualizada em",
    "Descrição completa",
    "Início Desc.",
    "Entregue Desc.",
    "Início Rev.",
    "Entregue Rev.",
    "Data Final",
    "1ª entrega do lote (Desc.)",
    "1ª entrega do lote (Rev.)",
]


class _EcoCSV:
    """Retorna o texto produzido pelo csv.writer sem reter o arquivo."""

    def write(self, texto):
        return texto


def _data_csv(valor, incluir_hora=False):
    """Replica o formato de datas da aba Geral do XLSX."""
    if not valor:
        return ""
    if hasattr(valor, "tzinfo") and valor.tzinfo is not None:
        if timezone.is_aware(valor):
            valor = timezone.localtime(valor)
    return valor.strftime("%d/%m/%Y %H:%M" if incluir_hora else "%d/%m/%Y")


def _celula_csv(valor):
    """Evita que Excel/LibreOffice interpretem conteúdo como fórmula."""
    if valor is None:
        return ""
    texto = str(valor)
    if texto.lstrip(" \t\r\n").startswith(("=", "+", "-", "@")):
        return "'" + texto
    return texto


def _imagens_em_blocos(queryset, tamanho=200):
    """Paginação por chave, sem carregar todas as imagens na RAM.

    Evita cursores server-side que podem não funcionar com o pooler do Neon.
    A retranca é única no modelo e desempata a ordenação por nome da obra.
    """
    queryset = queryset.order_by("nome_obra", "retranca")
    ultima_chave = None

    while True:
        pagina = queryset
        if ultima_chave is not None:
            ultima_obra, ultima_retranca = ultima_chave
            pagina = pagina.filter(
                Q(nome_obra__gt=ultima_obra)
                | Q(nome_obra=ultima_obra, retranca__gt=ultima_retranca)
            )

        bloco = list(pagina[:tamanho])
        if not bloco:
            break

        for imagem in bloco:
            yield imagem

        ultima = bloco[-1]
        ultima_chave = (ultima.nome_obra, ultima.retranca)


def _linhas_relatorio_csv(queryset, colunas):
    """Produz o cabeçalho e uma linha por imagem, sem montar arquivo inteiro."""
    escritor = csv.writer(_EcoCSV(), delimiter=";", lineterminator="\r\n")

    # BOM facilita a detecção de UTF-8 pelo Excel no Windows.
    yield "\ufeff"
    yield escritor.writerow(colunas)

    cache_lotes = {}
    for imagem in _imagens_em_blocos(queryset):
        descricao = getattr(imagem, "descricao", None)
        if descricao:
            # descricao__trechos já foi prefetched por _imagens_relatorio.
            # Filtrar via .filter() aqui geraria uma consulta extra por imagem.
            trechos = sorted(
                (trecho for trecho in descricao.trechos.all() if trecho.ativo),
                key=lambda trecho: trecho.ordem,
            )
            descricao_flat = " ".join(t.texto for t in trechos)
        else:
            descricao_flat = ""

        lote = imagem.lote
        responsavel = (
            imagem.responsavel.get_full_name() or imagem.responsavel.email
            if imagem.responsavel else ""
        )
        descritor = (
            descricao.descritor.get_full_name() or descricao.descritor.email
            if descricao and descricao.descritor else ""
        )
        revisor = (
            descricao.revisor.get_full_name() or descricao.revisor.email
            if descricao and descricao.revisor else ""
        )
        coordenador = (
            descricao.coordenador.get_full_name() or descricao.coordenador.email
            if descricao and descricao.coordenador else ""
        )

        valores = [
            imagem.retranca,
            imagem.nome_obra,
            imagem.componente_curricular.nome if imagem.componente_curricular else "",
            imagem.volume_ano_modulo,
            imagem.capitulo_unidade,
            imagem.get_etapa_display(),
            lote.nome if lote else "",
            _data_csv(lote.data_prevista if lote else None),
            _data_csv(lote.data_efetiva if lote else None),
            imagem.status.nome if imagem.status else "",
            responsavel,
            descritor,
            revisor,
            coordenador,
            imagem.get_pagamento_descritor_display(),
            imagem.get_pagamento_revisor_display(),
            _data_csv(imagem.prazo),
            _data_csv(imagem.criado_em, incluir_hora=True),
            _data_csv(imagem.atualizado_em, incluir_hora=True),
            descricao_flat or None,
        ]
        selecionados = selecionar_valores_dito(
            imagem, valores, colunas,
            datas=datas_formatadas_dito(imagem, _data_csv, cache_lotes),
        )
        yield escritor.writerow([_celula_csv(valor) for valor in selecionados])


@login_required
def relatorios_exportar_csv(request):
    """Exporta todas as imagens do resultado filtrado, sem limite de 200."""
    queryset = imagens_exportacao_dito(request)
    colunas = colunas_exportacao_dito()
    resposta = StreamingHttpResponse(
        _linhas_relatorio_csv(queryset, colunas),
        content_type="text/csv; charset=utf-8",
    )
    resposta["Content-Disposition"] = (
        'attachment; filename="relatorio_geral_dito.csv"'
    )
    resposta["X-Content-Type-Options"] = "nosniff"
    return resposta
