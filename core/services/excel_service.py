"""Leitura de valores em relatórios de importação (funções puras)."""

def _valor_excel(data, *chaves):
    """
    Retorna o primeiro valor preenchido entre os nomes de coluna aceitos.
    """
    for chave in chaves:
        valor = data.get(chave)

        if valor is not None and str(valor).strip():
            return valor

    return ""
