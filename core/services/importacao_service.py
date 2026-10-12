"""Análise em leitura de relatórios FotoWeb antes da importação.

Nenhuma função deste módulo utiliza banco de dados ou altera workflow.
"""
from collections import Counter, defaultdict
from pathlib import Path
import re
import unicodedata


def normalizar_cabecalho(valor):
    """Mesma estratégia de normalização dos cabeçalhos do importador web."""
    texto = unicodedata.normalize("NFKD", str(valor or ""))
    texto = texto.encode("ascii", "ignore").decode("ascii").casefold()
    return re.sub(r"[^a-z0-9]+", "_", texto).strip("_")


def normalizar_status(valor):
    """Agrupa variantes ortográficas; não decide status de destino no Dito!."""
    if valor is None:
        return "<VAZIO>"
    texto = unicodedata.normalize("NFKD", str(valor))
    texto = texto.encode("ascii", "ignore").decode("ascii").upper()
    texto = re.sub(r"\s+", " ", texto).strip()
    return texto if texto else "<VAZIO>"


def resumir_status_linhas(cabecalhos, linhas, max_amostras=3):
    """Retorna contagem de status por retranca em memória, sem persistência.

    Colunas necessárias: retranca e status. Duplicatas são contadas como
    linhas do relatório, não como imagens únicas do banco de dados.
    """
    nomes = [normalizar_cabecalho(c) for c in cabecalhos]
    if nomes.count("status") != 1 or nomes.count("retranca") != 1:
        raise ValueError("Relatório precisa ter exatamente uma coluna 'status' e uma 'retranca'.")
    i_status = nomes.index("status")
    i_retranca = nomes.index("retranca")
    contagem = Counter()
    amostras = defaultdict(list)
    total = sem_retranca = sem_status = 0
    for linha in linhas:
        total += 1
        retranca = str(linha[i_retranca] or "").strip() if i_retranca < len(linha) else ""
        status = normalizar_status(linha[i_status]) if i_status < len(linha) else "<VAZIO>"
        if not retranca:
            sem_retranca += 1
            continue
        if status == "<VAZIO>":
            sem_status += 1
        contagem[status] += 1
        if len(amostras[status]) < max_amostras:
            amostras[status].append(retranca)
    return {
        "total_linhas": total,
        "sem_retranca": sem_retranca,
        "sem_status": sem_status,
        "com_retranca": total - sem_retranca,
        "status": dict(contagem.most_common()),
        "amostras": dict(amostras),
    }


def auditar_arquivo_excel(caminho, max_amostras=3):
    """Abre .xlsx somente para leitura; não acessa Django nem banco."""
    import openpyxl

    arquivo = Path(caminho)
    if not arquivo.is_file() or arquivo.suffix.lower() != ".xlsx":
        raise ValueError("Informe um arquivo .xlsx existente.")
    wb = openpyxl.load_workbook(arquivo, read_only=True, data_only=True)
    try:
        ws = wb.active
        linhas = ws.iter_rows(values_only=True)
        cabecalhos = next(linhas, None)
        if cabecalhos is None:
            raise ValueError("O Excel não contém cabeçalhos.")
        return resumir_status_linhas(cabecalhos, linhas, max_amostras=max_amostras)
    finally:
        wb.close()


# DITO - Mapeamento dos status originais do FotoWeb
# Regras confirmadas em 12/10/2026.
# Funcoes puras: nao consultam nem modificam o banco.

_ALIASES_STATUS_FOTOWEB = {
    0: ("EXCLUIDA", "EXCLUIDAS"),
    1: ("DESCREVER", "LIBERADO PARA DESCRICAO", "LIBERADA PARA DESCRICAO"),
    2: ("DESCREVENDO", "EM DESCRICAO"),
    3: ("DESCRITA", "DESCRITAS", "DESCRITO", "DESCRITOS",
        "DESCRICAO CONCLUIDA"),
    4: ("CONFERINDO", "EM CONFERENCIA"),
    5: ("CONFERIDA", "CONFERIDAS", "CONFERIDO", "CONFERIDOS",
        "CONFERENCIA CONCLUIDA"),
    6: ("REVISANDO", "EM REVISAO"),
    7: ("REVISADA", "REVISADAS", "REVISADO", "REVISADOS"),
    8: ("COTEJANDO",),
    9: ("COTEJADA", "COTEJADAS"),
}

_DESTINOS_DITO = {
    1: "Liberado para descrição",
    2: "Descrevendo",
    3: "Descrito",
    4: "Conferindo",
    5: "Conferido",
    6: "Avaliando",
    7: "Avaliando",
}

_ALIAS_PARA_CODIGO = {
    normalizar_status(alias): codigo
    for codigo, nomes in _ALIASES_STATUS_FOTOWEB.items()
    for alias in nomes
}


def classificar_status_fotoweb(valor):
    """
    Retorna (acao, destino, motivo).

    Acoes: importar, ignorar ou revisar.
    O destino e o nome logico, ainda nao um objeto do banco.
    """
    texto = normalizar_status(valor)

    if texto == "<VAZIO>":
        return ("revisar", None, "Status vazio")

    encontrado = re.fullmatch(
        r"(?:(\d+)\s*[.:-]\s*)?(.+)",
        texto,
    )

    if not encontrado:
        return ("revisar", None, "Formato invalido")

    codigo_texto, nome = encontrado.groups()
    codigo_nome = _ALIAS_PARA_CODIGO.get(nome)

    if codigo_nome is None:
        return ("revisar", None, "Status desconhecido")

    if (
        codigo_texto is not None
        and int(codigo_texto) != codigo_nome
    ):
        return (
            "revisar",
            None,
            "Codigo e nome do status nao correspondem",
        )

    if codigo_nome in (0, 8, 9):
        return ("ignorar", None, "Excluida ou status legado")

    return (
        "importar",
        _DESTINOS_DITO[codigo_nome],
        "Reconhecido",
    )
