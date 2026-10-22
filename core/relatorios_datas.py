"""Datas exibidas e exportadas no Dito!.

Consulta histórico por IMAGEM, inclusive eventos antigos sem vínculo com
Descricao. Os eventos são recuperados em lotes, sem N+1 queries.
Não inventa datas: etapas sem evento permanecem vazias.
"""

from types import SimpleNamespace

from django.db.models import Min, Prefetch, Q
from .models import HistoricoItem, Imagem


_TIPOS = (
    HistoricoItem.TipoAcao.DESCRICAO_INICIADA,
    HistoricoItem.TipoAcao.DESCRICAO_SALVA,
    HistoricoItem.TipoAcao.CONFERENCIA_INICIADA,
    HistoricoItem.TipoAcao.CONFERENCIA_CONCLUIDA,
)


def _eventos_datas_query():
    """Datas editoriais e entrada no status final, inclusive alteracao manual.

    Nao usa imagem.prazo ou imagem.atualizado_em: ambos podem divergir
    da data efetiva de conclusao. Consulta o proprio historico persistido.
    """
    return (
        HistoricoItem.objects
        .filter(
            Q(tipo_acao__in=_TIPOS)
            | Q(tipo_acao=HistoricoItem.TipoAcao.DESCRICAO_FINALIZADA)
            | Q(novo_status__is_final=True)
        )
        .select_related("novo_status")
        .only(
            "id", "imagem_id", "tipo_acao", "criado_em",
            "novo_status_id", "novo_status__is_final",
        )
        .order_by("criado_em", "id")
    )


def historico_prefetch_dito():
    """Prefetch único reutilizado pela tabela, XLSX e CSV."""
    eventos = _eventos_datas_query()
    return Prefetch(
        "historico", queryset=eventos, to_attr="_eventos_exportacao_dito"
    )


def imagens_exportacao_dito(request):
    # Import local evita referência circular com core.views.
    from .views import _imagens_relatorio

    return _imagens_relatorio(
        request, aplicar_filtros=True
    ).prefetch_related(historico_prefetch_dito())


def datas_imagem_dito(imagem):
    """Início Desc, Entregue Desc, Início Rev, Entregue Rev, Data Final.

    Usa historico por imagem, inclusive eventos sem descricao_id.
    Data Final e a ultima entrada documentada em um status final (ou
    evento de finalizacao); NAO e o prazo previsto da imagem.
    """
    eventos = getattr(imagem, "_eventos_exportacao_dito", None)
    if eventos is None:
        eventos = _eventos_datas_query().filter(imagem_id=imagem.pk)

    inicio_desc = fim_desc = inicio_rev = fim_rev = data_final = None

    for evento in eventos:
        acao = evento.tipo_acao
        if acao == HistoricoItem.TipoAcao.DESCRICAO_INICIADA and inicio_desc is None:
            inicio_desc = evento.criado_em
        elif acao == HistoricoItem.TipoAcao.DESCRICAO_SALVA:
            fim_desc = evento.criado_em  # Último salvamento registrado
        elif acao == HistoricoItem.TipoAcao.CONFERENCIA_INICIADA and inicio_rev is None:
            inicio_rev = evento.criado_em
        elif acao == HistoricoItem.TipoAcao.CONFERENCIA_CONCLUIDA:
            fim_rev = evento.criado_em  # Última entrega; a primeira permanece no histórico

        # Captura tambem avancos cujo tipo_acao nao seja "descricao_finalizada"
        # (ex.: revisao_concluida ou STATUS_ALTERADO para um status final).
        # A ordenacao ascendente preserva a ultima conclusao registrada.
        if (
            acao == HistoricoItem.TipoAcao.DESCRICAO_FINALIZADA
            or (
                evento.novo_status_id is not None
                and evento.novo_status is not None
                and evento.novo_status.is_final
            )
        ):
            data_final = evento.criado_em

    return inicio_desc, fim_desc, inicio_rev, fim_rev, data_final


def primeiras_entregas_lote_dito(imagem, cache=None):
    """Primeira data em que TODAS as imagens ativas do lote foram entregues.

    O marco e o ultimo entre os PRIMEIROS eventos de conclusao de cada
    imagem. Correcao/reentrega nao reescreve esse marco. Os registros sao
    os HistoricoItem ja persistidos; nenhuma migracao e necessaria.

    Se faltar evento de conclusao para uma imagem, nao inventamos o marco.
    Se a composicao de um lote mudar depois, o marco historico nao pode
    ser garantido sem um registro proprio e imutavel do lote.
    """
    lote_id = getattr(imagem, "lote_id", None)
    if lote_id is None:
        return None, None
    if cache is None:
        cache = {}
    if lote_id not in cache:
        quantidade = Imagem.objects.filter(lote_id=lote_id, ativo=True).count()
        primeiros_por_tipo = {
            HistoricoItem.TipoAcao.DESCRICAO_SALVA: {},
            HistoricoItem.TipoAcao.CONFERENCIA_CONCLUIDA: {},
        }
        if quantidade:
            registros = (
                HistoricoItem.objects.filter(
                    imagem__lote_id=lote_id,
                    imagem__ativo=True,
                    tipo_acao__in=tuple(primeiros_por_tipo),
                )
                .order_by()
                .values("imagem_id", "tipo_acao")
                .annotate(primeira=Min("criado_em"))
                .values_list("imagem_id", "tipo_acao", "primeira")
            )
            for imagem_id, tipo, primeira in registros:
                primeiros_por_tipo[tipo][imagem_id] = primeira

        def primeira_entrega_completa(tipo):
            datas = primeiros_por_tipo[tipo]
            return max(datas.values()) if quantidade and len(datas) == quantidade else None

        cache[lote_id] = (
            primeira_entrega_completa(HistoricoItem.TipoAcao.DESCRICAO_SALVA),
            primeira_entrega_completa(HistoricoItem.TipoAcao.CONFERENCIA_CONCLUIDA),
        )
    return cache[lote_id]


def datas_tela_dito(imagem, cache_lotes=None):
    valores = datas_imagem_dito(imagem)
    lote_desc, lote_rev = primeiras_entregas_lote_dito(imagem, cache_lotes)
    return SimpleNamespace(
        inicio_desc=valores[0],
        entregue_desc=valores[1],
        inicio_rev=valores[2],
        entregue_rev=valores[3],
        data_final=valores[4],
        primeira_entrega_lote_desc=lote_desc,
        primeira_entrega_lote_rev=lote_rev,
    )


def datas_formatadas_dito(imagem, formatar, cache_lotes=None):
    valores = datas_imagem_dito(imagem)
    lote_desc, lote_rev = primeiras_entregas_lote_dito(imagem, cache_lotes)
    return [
        formatar(valores[0], incluir_hora=True),
        formatar(valores[1], incluir_hora=True),
        formatar(valores[2], incluir_hora=True),
        formatar(valores[3], incluir_hora=True),
        formatar(valores[4]),
        formatar(lote_desc, incluir_hora=True),
        formatar(lote_rev, incluir_hora=True),
    ]
