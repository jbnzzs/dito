"""Percurso persistente por usuário e lote (DITO_PERCURSO_PERSISTENTE_V2).

A ordem de trabalho não é recalculada a cada avanço: apenas tarefas novas
são acrescentadas no fim. O histórico registra visitas, não conclusões.
Abrir uma imagem nunca altera a flag pronto_para_lote.
"""

MAX_HISTORICO = 5000


def _inteiros_unicos(valores):
    resultado = []
    vistos = set()
    for valor in valores or []:
        try:
            numero = int(valor)
        except (TypeError, ValueError):
            continue
        if numero > 0 and numero not in vistos:
            resultado.append(numero)
            vistos.add(numero)
    return resultado


def _sincronizar_ordem(ordem, novas, atual=None):
    """Mantém as posições existentes; apenas acrescenta IDs novos."""
    resultado = _inteiros_unicos(ordem)
    presentes = set(resultado)
    for numero in _inteiros_unicos(novas):
        if numero not in presentes:
            resultado.append(numero)
            presentes.add(numero)
    if atual is not None and int(atual) not in presentes:
        resultado.append(int(atual))
    return resultado


def _selecionar_proxima(ordem, atual_id, pendentes, historico=(), indice=0, autorizadas=None):
    """Se houve Voltar, refaz o caminho real; senão, usa ordem estável.

    autorizadas inclui imagens marcadas prontas, pois podem precisar ser
    revisitadas após o botão Voltar. pendentes exclui imagens prontas.
    """
    atual_id = int(atual_id)
    pendentes = set(_inteiros_unicos(pendentes))
    autorizadas = set(_inteiros_unicos(autorizadas if autorizadas is not None else pendentes))
    historico = [int(x) for x in historico or [] if str(x).isdigit()]
    indice = int(indice or 0)
    if 0 <= indice < len(historico) and historico[indice] == atual_id:
        if indice + 1 < len(historico):
            seguinte = historico[indice + 1]
            if seguinte != atual_id and seguinte in autorizadas:
                return seguinte

    ordem = _inteiros_unicos(ordem)
    if atual_id in ordem:
        indice_atual = ordem.index(atual_id)
        candidatos = ordem[indice_atual + 1:] + ordem[:indice_atual]
    else:
        candidatos = ordem

    for imagem_id in candidatos:
        if imagem_id != atual_id and imagem_id in pendentes:
            return imagem_id
    return None


def registrar_visita_lote(request, lote_id, imagem_id, total_disponivel, ids_pendentes):
    """Guarda cursor e histórico no banco; retorna (passo, total, anterior_id)."""
    from django.db import transaction
    from core.models import PercursoLoteUsuario

    lote_id = int(lote_id)
    imagem_id = int(imagem_id)
    ordenacao_inicial = _sincronizar_ordem(ids_pendentes, [], imagem_id)
    # Recupera o histórico V1 quando ainda existir na sessão atual. A partir
    # daqui a informação passa a ser persistida no banco por usuário/lote.
    legado = request.session.get("dito_percurso_usuario_v1", {}) if hasattr(request, "session") else {}
    passos_iniciais = [imagem_id]
    indice_inicial = 0
    if (
        isinstance(legado, dict)
        and legado.get("usuario_id") == request.user.pk
        and legado.get("lote_id") == lote_id
        and isinstance(legado.get("passos"), list)
    ):
        passos_legado = [int(x) for x in legado["passos"] if str(x).isdigit()]
        if passos_legado:
            removidos = max(0, len(passos_legado) - MAX_HISTORICO)
            passos_iniciais = passos_legado[removidos:]
            indice_bruto = legado.get("indice", 0)
            try:
                indice_bruto = int(indice_bruto)
            except (TypeError, ValueError):
                indice_bruto = 0
            indice_inicial = min(max(0, indice_bruto - removidos), len(passos_iniciais) - 1)
            ordenacao_inicial = _sincronizar_ordem(ordenacao_inicial, passos_iniciais)

    with transaction.atomic():
        estado, _ = (
            PercursoLoteUsuario.objects.select_for_update().get_or_create(
                usuario=request.user,
                lote_id=lote_id,
                defaults={
                    "ordem_ids": ordenacao_inicial,
                    "historico_ids": passos_iniciais,
                    "indice": indice_inicial,
                    "total_inicial": max(1, int(total_disponivel), len(ordenacao_inicial)),
                },
            )
        )
        ordem = _sincronizar_ordem(estado.ordem_ids, ids_pendentes, imagem_id)
        passos = [int(x) for x in estado.historico_ids if str(x).isdigit()]
        indice = min(max(0, estado.indice), len(passos) - 1) if passos else 0
        modo = request.GET.get("dito_percurso")

        if not passos:
            passos = [imagem_id]
            indice = 0
        elif passos[indice] == imagem_id:
            pass  # Refresh ou retorno à imagem atual: nada a acrescentar.
        elif modo == "anterior" and indice > 0 and passos[indice - 1] == imagem_id:
            indice -= 1
        elif modo == "proximo" and indice + 1 < len(passos) and passos[indice + 1] == imagem_id:
            indice += 1
        else:
            # Abertura direta ou troca de direção: preserva o passado real.
            passos = passos[:indice + 1] + [imagem_id]
            indice = len(passos) - 1

        if len(passos) > MAX_HISTORICO:
            removidos = len(passos) - MAX_HISTORICO
            passos = passos[removidos:]
            indice = max(0, indice - removidos)

        estado.ordem_ids = ordem
        estado.historico_ids = passos
        estado.indice = indice
        estado.total_inicial = max(1, int(estado.total_inicial), int(total_disponivel), len(ordem))
        estado.save(update_fields=["ordem_ids", "historico_ids", "indice", "total_inicial", "atualizado_em"])

    anterior_id = passos[indice - 1] if indice > 0 else None
    passo = min(len(set(passos[:indice + 1])), estado.total_inicial)
    return passo, estado.total_inicial, anterior_id


def proxima_imagem_percurso(usuario, lote_id, imagem_id, ids_pendentes, ids_autorizados):
    """Retorna ID da próxima imagem, sem alterar a posição do histórico."""
    from core.models import PercursoLoteUsuario

    estado = PercursoLoteUsuario.objects.filter(usuario=usuario, lote_id=lote_id).first()
    if estado:
        ordem = _sincronizar_ordem(estado.ordem_ids, ids_pendentes, imagem_id)
        return _selecionar_proxima(
            ordem, imagem_id, ids_pendentes, estado.historico_ids,
            estado.indice, ids_autorizados,
        )
    return _selecionar_proxima(
        _sincronizar_ordem(ids_pendentes, [], imagem_id),
        imagem_id, ids_pendentes, autorizadas=ids_autorizados,
    )


def imagem_para_retomar(usuario, lote_id, ids_pendentes, ids_autorizados):
    """Prefere o cursor ainda autorizado; caso contrário, primeira pendente."""
    from core.models import PercursoLoteUsuario

    autorizados = set(_inteiros_unicos(ids_autorizados))
    pendentes = set(_inteiros_unicos(ids_pendentes))
    estado = PercursoLoteUsuario.objects.filter(usuario=usuario, lote_id=lote_id).first()
    if estado:
        passos = estado.historico_ids or []
        ordem = _sincronizar_ordem(estado.ordem_ids, ids_pendentes)
        if passos:
            indice = min(max(0, estado.indice), len(passos) - 1)
            atual = passos[indice]
            if atual in autorizados and atual in pendentes:
                return atual  # Interrompida e ainda não concluída.
            if atual in autorizados and pendentes:
                proxima = _selecionar_proxima(ordem, atual, pendentes)
                if proxima is not None:
                    return proxima  # Última já pronta: passa à próxima pendente.
    else:
        ordem = _inteiros_unicos(ids_pendentes)

    for imagem_id in ordem:
        if imagem_id in pendentes and imagem_id in autorizados:
            return imagem_id
    return None
