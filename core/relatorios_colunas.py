"""Colunas dinamicas do relatorio Dito (NAO se aplica ao FotoWeb).

Sem filtros, exporta as 20 colunas originais; com filtros, Retranca
e os campos correspondentes exatamente aos filtros selecionados. Os filtros continuam responsaveis
por determinar QUAIS imagens sao retornadas em _imagens_relatorio(request).

A lista em _NOMES_ORIGINAIS espelha as 20 colunas da aba Geral existente
antes desta melhoria. As colunas fixas do FotoWeb nao sao modificadas.
"""

_NOMES_ORIGINAIS = (
    "Retranca",
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
)

# Ordem segue a interface dos filtros, com Retranca sempre presente.
MAPA_FILTROS = (
    ("projeto", "Projeto"),
    ("obra", "Obra"),
    ("componente", "Componente curricular"),
    ("lote", "Lote"),
    ("status", "Status"),
    ("responsavel", "Responsável atual"),
    ("pagamento_descritor", "Pagamento descritor"),
    ("pagamento_revisor", "Pagamento revisor"),
    ("data_inicio", "Cadastrada em"),
    ("data_fim", "Cadastrada em"),
)


def colunas_dito(request):
    """Todas as colunas sem filtros; caso contrario, Retranca + filtros.

    A restricao e apenas de COLUNAS. O total de LINHAS continua vindo do
    mesmo queryset compartilhado pela tabela e pelas exportacoes.
    """
    filtros = ("retranca",) + tuple(nome for nome, _ in MAPA_FILTROS)
    if not any(str(request.GET.get(nome, "") or "").strip() for nome in filtros):
        return list(_NOMES_ORIGINAIS)

    colunas = ["Retranca"]
    for nome_filtro, titulo_coluna in MAPA_FILTROS:
        valor = request.GET.get(nome_filtro, "")
        if str(valor or "").strip() and titulo_coluna not in colunas:
            colunas.append(titulo_coluna)
    return colunas


def selecionar_valores_dito(imagem, valores_originais, colunas):
    """Seleciona os valores nas mesmas colunas/ordem da tabela e do CSV.

    `valores_originais` sao os 20 valores da aba Geral legada, ja calculados
    pelo exportador. FotoWeb continua independente.
    """
    if len(valores_originais) != len(_NOMES_ORIGINAIS):
        raise ValueError("Numero inesperado de colunas na exportacao Dito.")
    mapa = dict(zip(_NOMES_ORIGINAIS, valores_originais))
    projeto = getattr(imagem, "projeto", None)
    mapa["Projeto"] = projeto.nome if projeto else ""
    return [mapa[nome] for nome in colunas]


def larguras_dito(colunas, larguras_originais):
    if len(larguras_originais) != len(_NOMES_ORIGINAIS):
        raise ValueError("Numero inesperado de larguras no Excel Dito.")
    mapa = dict(zip(_NOMES_ORIGINAIS, larguras_originais))
    mapa["Projeto"] = 26
    return [mapa[nome] for nome in colunas]
