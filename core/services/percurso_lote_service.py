"""Percurso das imagens do lote por sessão e usuário.

A numeração representa a ordem real de visita durante esta entrada no lote,
independentemente da retranca, do status ou do perfil atual. Não grava dados
no banco e não altera a navegação e as permissões do workflow.
"""

CHAVE = "dito_percurso_usuario_v1"
MAX_HISTORICO = 1000  # Evita crescimento ilimitado ao circular pelo mesmo lote.


def registrar_visita_lote(request, lote_id, imagem_id, total_disponivel):
    """Retorna (passo, total, imagem_anterior_id) da visita autorizada.

    Sem marcador na URL: uma entrada nova começa em 1.
    dito_percurso=proximo: avança uma posição, sem contar recarregamentos.
    dito_percurso=anterior: volta à posição visitada anteriormente.
    """
    usuario_id = request.user.pk
    lote_id, imagem_id = int(lote_id), int(imagem_id)
    modo = request.GET.get("dito_percurso")
    estado = request.session.get(CHAVE)

    valido = (
        isinstance(estado, dict)
        and estado.get("usuario_id") == usuario_id
        and estado.get("lote_id") == lote_id
        and isinstance(estado.get("passos"), list)
        and bool(estado["passos"])
        and isinstance(estado.get("indice"), int)
        and 0 <= estado["indice"] < len(estado["passos"])
    )

    if not valido or modo not in ("proximo", "anterior"):
        estado = {
            "usuario_id": usuario_id,
            "lote_id": lote_id,
            "passos": [imagem_id],
            "indice": 0,
            "inicio": 0,
            "total": max(1, int(total_disponivel)),
        }
    else:
        passos = list(estado["passos"])
        indice = estado["indice"]
        if modo == "proximo":
            if passos[indice] == imagem_id:
                pass  # Atualização da mesma página: não aumenta o contador.
            elif indice + 1 < len(passos) and passos[indice + 1] == imagem_id:
                indice += 1
            else:
                # Uma nova direção depois de voltar substitui apenas o futuro.
                passos = passos[:indice + 1]
                passos.append(imagem_id)
                indice += 1
        else:  # anterior
            if indice > 0 and passos[indice - 1] == imagem_id:
                indice -= 1
            elif passos[indice] != imagem_id:
                # Link antigo ou histórico diferente: inicia novo percurso.
                passos = [imagem_id]
                indice = 0
                estado["inicio"] = 0
        inicio = estado.get("inicio", 0)
        if len(passos) > MAX_HISTORICO:
            remover = len(passos) - MAX_HISTORICO
            passos = passos[remover:]
            indice -= remover
            inicio += remover
        estado = {**estado, "passos": passos, "indice": indice, "inicio": inicio}

    request.session[CHAVE] = estado
    indice = estado["indice"]
    anterior_id = estado["passos"][indice - 1] if indice > 0 else None
    return estado.get("inicio", 0) + indice + 1, max(1, int(estado["total"])), anterior_id
