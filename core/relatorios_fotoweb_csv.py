"""CSV FotoWeb do Dito! - exportação em lotes e compatível com o layout XLSX.

Mantém as 13 colunas, os valores dos snapshots originais, o agrupamento por
arquivo/linha da importação original e o comportamento do XLSX existente.

Importante: o CSV ainda precisa ser testado no importador FotoWeb da empresa.
"""

import csv
import json
from datetime import date, datetime
from decimal import Decimal

from django.contrib import messages
from django.contrib.auth.decorators import login_required
from django.db.models import BigIntegerField, CharField, Q, Value
from django.db.models.fields.json import KeyTextTransform
from django.db.models.functions import Cast, Coalesce
from django.http import StreamingHttpResponse
from django.shortcuts import redirect
from django.urls import reverse
from django.utils import timezone


TAMANHO_BLOCO = 200
VALOR_LINHA_AUSENTE = 10**12


class _Echo:
    def write(self, value):
        return value


def _celula(valor):
    if valor is None:
        return ""
    if isinstance(valor, bool):
        return "Sim" if valor else "Não"
    if isinstance(valor, Decimal):
        return format(valor, "f").replace(".", ",")
    if isinstance(valor, float):
        return str(valor).replace(".", ",")
    if isinstance(valor, datetime):
        if timezone.is_aware(valor):
            valor = timezone.localtime(valor)
        return valor.strftime("%d/%m/%Y %H:%M")
    if isinstance(valor, date):
        return valor.strftime("%d/%m/%Y")
    texto = str(valor)
    # Proteção para Excel/Sheets. Apenas texto recebe apóstrofo.
    if texto.lstrip(" \t\r\n").startswith(("=", "+", "-", "@")):
        return "'" + texto
    return texto


def _ordenar_fotoweb(queryset):
    """Ordena no BD pelo arquivo de origem, número da linha e retranca.

    Essas chaves existem nos snapshots gerados pelo importador do Dito!.
    A ordenação substitui `list(...); imagens.sort(...)` em memória.
    """
    return (
        queryset.annotate(
            _csv_arquivo=Coalesce(
                KeyTextTransform("__arquivo_origem", "dados_fotoweb_originais"),
                Value(""),
                output_field=CharField(),
            ),
            _csv_linha=Coalesce(
                Cast(
                    KeyTextTransform("__numero_linha", "dados_fotoweb_originais"),
                    BigIntegerField(),
                ),
                Value(VALOR_LINHA_AUSENTE),
                output_field=BigIntegerField(),
            ),
        )
        .order_by("_csv_arquivo", "_csv_linha", "retranca")
    )


def _imagens_por_blocos(queryset):
    """Paginação por chave composta, sem cursores server-side do PostgreSQL."""
    queryset = _ordenar_fotoweb(queryset)
    chave = None
    while True:
        pagina = queryset
        if chave is not None:
            arquivo, numero, retranca = chave
            pagina = pagina.filter(
                Q(_csv_arquivo__gt=arquivo)
                | Q(_csv_arquivo=arquivo, _csv_linha__gt=numero)
                | Q(
                    _csv_arquivo=arquivo,
                    _csv_linha=numero,
                    retranca__gt=retranca,
                )
            )
        bloco = list(pagina[:TAMANHO_BLOCO])
        if not bloco:
            break
        for imagem in bloco:
            yield imagem
        ultima = bloco[-1]
        chave = (ultima._csv_arquivo, ultima._csv_linha, ultima.retranca)


def _conteudo_csv(queryset):
    from .views import FOTOWEB_COLUNAS_RELATORIO, _lang_tag

    writer = csv.writer(_Echo(), delimiter=";", lineterminator="\r\n")
    yield "\ufeff"
    yield writer.writerow(FOTOWEB_COLUNAS_RELATORIO)

    for imagem in _imagens_por_blocos(queryset):
        base = dict(imagem.dados_fotoweb_originais or {})
        descricao = getattr(imagem, "descricao", None)
        trechos = []
        if descricao:
            # Aproveita o prefetch realizado pela consulta base em cada bloco.
            trechos = sorted(
                (t for t in descricao.trechos.all() if t.ativo),
                key=lambda t: t.ordem,
            )

        base["obra"] = imagem.nome_obra
        base["componente"] = (
            imagem.componente_curricular.nome if imagem.componente_curricular else ""
        )
        base["volume"] = imagem.volume_ano_modulo or ""
        base["capitulo"] = imagem.capitulo_unidade or ""
        base["status"] = imagem.status.nome if imagem.status else ""
        base["retranca"] = imagem.retranca
        base["retranca_lower"] = imagem.retranca.lower() if imagem.retranca else ""

        usuario = None
        if descricao and descricao.descritor:
            usuario = descricao.descritor
        elif imagem.responsavel:
            usuario = imagem.responsavel
        if usuario:
            base["usuario"] = usuario.username or usuario.email or ""

        base["etapa"] = f"Etapa: {imagem.etapa}" if imagem.etapa else ""
        if trechos:
            base["descricao"] = json.dumps(
                [
                    {"lang": _lang_tag(t.idioma_codigo), "text": t.texto}
                    for t in trechos
                ],
                ensure_ascii=False,
            )
            base["descricao_flat"] = " ".join(t.texto for t in trechos).strip()
        else:
            base["descricao"] = None
            base["descricao_flat"] = None

        yield writer.writerow([_celula(base.get(c)) for c in FOTOWEB_COLUNAS_RELATORIO])


def _voltar_com_filtros(request):
    destino = reverse("relatorios_lista")
    query = request.GET.urlencode()
    return redirect(f"{destino}?{query}" if query else destino)


@login_required
def relatorios_exportar_fotoweb_csv(request):
    from .views import _imagens_relatorio

    queryset = _imagens_relatorio(request)
    if not queryset.exists():
        messages.error(request, "Nenhuma imagem encontrada para exportar.")
        return _voltar_com_filtros(request)

    # O XLSX FotoWeb já exige o snapshot original; mantemos a mesma regra.
    sem_base = queryset.filter(
        Q(dados_fotoweb_originais={})
        | Q(dados_fotoweb_originais__isnull=True)
    )
    total_sem_base = sem_base.count()
    if total_sem_base:
        exemplos = ", ".join(sem_base.values_list("retranca", flat=True)[:5])
        messages.error(
            request,
            f"{total_sem_base} imagem(ns) não possuem a linha original do FotoWeb. "
            + (f"Exemplos: {exemplos}. " if exemplos else "")
            + "Reimporte o arquivo original para sincronizar a base FotoWeb.",
        )
        return _voltar_com_filtros(request)

    response = StreamingHttpResponse(
        _conteudo_csv(queryset), content_type="text/csv; charset=utf-8"
    )
    response["Content-Disposition"] = 'attachment; filename="relatorio_fotoweb_atualizado.csv"'
    response["X-Content-Type-Options"] = "nosniff"
    return response
