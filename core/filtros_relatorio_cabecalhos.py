"""Filtros por coluna do Relatorio Geral (servidor, nao somente pagina atual).

A whitelist abaixo mapeia as 27 colunas visuais para campos reais ou
marcos historicos. Filtros de datas trabalham em datas locais inclusivas.
Nao modifica registros nem cria novos campos no banco.
"""
from datetime import date

from django.db.models import Count, Min, OuterRef, Q, Subquery
from django.utils import timezone

from .models import HistoricoItem, Imagem


# As chaves espelham data-coluna no template; somente estas sao aceitas.
TEXTOS = {
    "retranca": "retranca",
    "volume": "volume_ano_modulo",
    "capitulo": "capitulo_unidade",
    "obra": "nome_obra",
    "componente": "componente_curricular__nome",
    "descricao": "descricao__trechos__texto",
}
ESCOLHAS = {
    "etapa": "etapa",
    "pag-descritor": "pagamento_descritor",
    "pag-revisor": "pagamento_revisor",
}
RELACOES_ID = {
    "projeto": ("projeto_id", "sem_projeto"),
    "lote": ("lote_id", "avulsas"),
    "status": ("status_id", "sem_status"),
    "responsavel": ("responsavel_id", "sem_responsavel"),
}
PESSOAS = {
    "descritor": "descricao__descritor",
    "revisor": "descricao__revisor",
    "coordenador": "descricao__coordenador",
}
DATAS_CAMPOS = {
    "prevista": ("lote__data_prevista", False),
    "efetiva": ("lote__data_efetiva", False),
    "prazo": ("prazo", False),
    "cadastrada": ("criado_em", True),
    "atualizada": ("atualizado_em", True),
}
DATAS_EVENTOS = {
    "inicio-desc": (HistoricoItem.TipoAcao.DESCRICAO_INICIADA, False),
    "entregue-desc": (HistoricoItem.TipoAcao.DESCRICAO_SALVA, True),
    "inicio-rev": (HistoricoItem.TipoAcao.CONFERENCIA_INICIADA, False),
    "entregue-rev": (HistoricoItem.TipoAcao.CONFERENCIA_CONCLUIDA, True),
    "data-final": ("__final__", True),
}
DATAS_LOTE = {
    "primeira-lote-desc": HistoricoItem.TipoAcao.DESCRICAO_SALVA,
    "primeira-lote-rev": HistoricoItem.TipoAcao.CONFERENCIA_CONCLUIDA,
}
TODAS_COLUNAS = tuple([
    *TEXTOS, *ESCOLHAS, *RELACOES_ID, *PESSOAS,
    *DATAS_CAMPOS, *DATAS_EVENTOS, *DATAS_LOTE,
])


def parametros_cabecalho(request):
    """Pares validos que precisam sobreviver ao envio do formulario antigo."""
    resultado = []
    for chave in TODAS_COLUNAS:
        campos = (f"cf_{chave}_de", f"cf_{chave}_ate") if (
            chave in DATAS_CAMPOS or chave in DATAS_EVENTOS or chave in DATAS_LOTE
        ) else (f"cf_{chave}",)
        for campo in campos:
            valor = (request.GET.get(campo) or "").strip()
            if valor:
                resultado.append({"nome": campo, "valor": valor})
    return resultado


def _data_iso(valor):
    if not valor:
        return None
    try:
        return date.fromisoformat(str(valor).strip())
    except ValueError:
        return None


def _filtrar_datas(qs, campo, de, ate, datetime_field=False):
    if de:
        qs = qs.filter(**{f"{campo}__date__gte" if datetime_field else f"{campo}__gte": de})
    if ate:
        qs = qs.filter(**{f"{campo}__date__lte" if datetime_field else f"{campo}__lte": ate})
    return qs


def _marcos_lotes(tipo_acao, de, ate):
    """IDs de lotes com primeira entrega completa na janela solicitada.

    Reproduz a regra atual de primeiras_entregas_lote_dito: maximo das
    primeiras entregas por imagem ativa. Nao assume datas sem evento.
    """
    totais = dict(
        Imagem.objects.filter(ativo=True, lote_id__isnull=False)
        .order_by().values("lote_id")
        .annotate(total=Count("pk"))
        .values_list("lote_id", "total")
    )
    primeiras = (
        HistoricoItem.objects.filter(
            tipo_acao=tipo_acao, imagem__ativo=True,
            imagem__lote_id__isnull=False,
        )
        .order_by()
        .values("imagem__lote_id", "imagem_id")
        .annotate(data=Min("criado_em"))
        .values_list("imagem__lote_id", "imagem_id", "data")
    )
    dados = {}
    for lote_id, imagem_id, data in primeiras:
        item = dados.setdefault(lote_id, [0, None])
        item[0] += 1
        if item[1] is None or data > item[1]:
            item[1] = data

    resultado = []
    for lote_id, (quantidade, marco) in dados.items():
        if quantidade != totais.get(lote_id) or marco is None:
            continue
        if timezone.is_aware(marco):
            marco = timezone.localtime(marco)
        dia = marco.date()
        if (de is None or dia >= de) and (ate is None or dia <= ate):
            resultado.append(lote_id)
    return resultado


def aplicar_filtros_cabecalho(qs, request):
    """Aplica combinacao AND dos filtros, antes da paginacao/exportacao."""
    for chave, campo in TEXTOS.items():
        valor = (request.GET.get(f"cf_{chave}") or "").strip()
        if valor:
            if chave == "descricao":
                qs = qs.filter(
                    descricao__trechos__ativo=True,
                    descricao__trechos__texto__icontains=valor,
                ).distinct()
            else:
                qs = qs.filter(**{f"{campo}__icontains": valor})

    for chave, campo in ESCOLHAS.items():
        valor = (request.GET.get(f"cf_{chave}") or "").strip()
        if valor:
            qs = qs.filter(**{campo: valor})

    for chave, (campo, valor_vazio) in RELACOES_ID.items():
        valor = (request.GET.get(f"cf_{chave}") or "").strip()
        if valor == valor_vazio:
            qs = qs.filter(**{f"{campo}__isnull": True})
        elif valor:
            if valor.isdecimal():
                qs = qs.filter(**{campo: int(valor)})
            else:
                qs = qs.none()

    for chave, campo in PESSOAS.items():
        valor = (request.GET.get(f"cf_{chave}") or "").strip()
        if valor:
            for termo in valor.split():
                qs = qs.filter(
                    Q(**{f"{campo}__first_name__icontains": termo})
                    | Q(**{f"{campo}__last_name__icontains": termo})
                    | Q(**{f"{campo}__email__icontains": termo})
                    | Q(**{f"{campo}__username__icontains": termo})
                )

    for chave, (campo, eh_datetime) in DATAS_CAMPOS.items():
        de = _data_iso(request.GET.get(f"cf_{chave}_de"))
        ate = _data_iso(request.GET.get(f"cf_{chave}_ate"))
        if de or ate:
            qs = _filtrar_datas(qs, campo, de, ate, eh_datetime)

    for chave, (tipo, ultimo) in DATAS_EVENTOS.items():
        de = _data_iso(request.GET.get(f"cf_{chave}_de"))
        ate = _data_iso(request.GET.get(f"cf_{chave}_ate"))
        if not (de or ate):
            continue
        historico = HistoricoItem.objects.filter(imagem_id=OuterRef("pk"))
        if tipo == "__final__":
            historico = historico.filter(
                Q(tipo_acao=HistoricoItem.TipoAcao.DESCRICAO_FINALIZADA)
                | Q(novo_status__is_final=True)
            )
        else:
            historico = historico.filter(tipo_acao=tipo)
        ordem = ("-criado_em", "-id") if ultimo else ("criado_em", "id")
        historico = historico.order_by(*ordem).values("criado_em")[:1]
        # alias dispensa coluna extra retornada pelo SQL; mesma semantica
        # da data mostrada na tabela (primeiro ou ultimo evento).
        qs = qs.alias(**{f"_dito_dt_{chave.replace('-', '_')}": Subquery(historico)})
        qs = _filtrar_datas(
            qs, f"_dito_dt_{chave.replace('-', '_')}", de, ate, True
        )

    for chave, tipo in DATAS_LOTE.items():
        de = _data_iso(request.GET.get(f"cf_{chave}_de"))
        ate = _data_iso(request.GET.get(f"cf_{chave}_ate"))
        if de or ate:
            qs = qs.filter(lote_id__in=_marcos_lotes(tipo, de, ate))

    return qs
