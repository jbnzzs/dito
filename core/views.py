import ast
import os

import openpyxl
import pycountry
from django.contrib import messages
from django.contrib.auth.decorators import login_required
from django.db import transaction
from django.shortcuts import redirect, render, get_object_or_404

from .models import Descricao, Imagem, StatusWorkflow, Trecho, Usuario
from .forms import ImagemForm


def _saudacao():
    from datetime import datetime
    hora = datetime.now().hour
    if hora < 12:
        return "Bom dia"
    elif hora < 18:
        return "Boa tarde"
    return "Boa noite"


def _apenas_coordenador(usuario):
    return usuario.tipo in (usuario.Tipo.ADMINISTRADOR, usuario.Tipo.COORDENADOR)


# ============================================================
# DASHBOARD
# ============================================================

@login_required
def dashboard(request):
    from datetime import datetime
    from django.db.models import Count, Q

    usuario = request.user
    hoje = datetime.now().strftime("%d/%m/%Y")
    ctx = {"saudacao": _saudacao(), "hoje": hoje}

    if usuario.tipo in (usuario.Tipo.ADMINISTRADOR, usuario.Tipo.COORDENADOR):
        total = Imagem.objects.filter(ativo=True).count()
        finalizadas = Imagem.objects.filter(ativo=True, status__slug="finalizado").count()
        pendentes = total - finalizadas
        por_status = (
            StatusWorkflow.objects
            .filter(ativo=True)
            .annotate(total=Count("imagens", filter=Q(imagens__ativo=True)))
            .order_by("ordem")
        )
        from .models import HistoricoItem
        historico_recente = (
            HistoricoItem.objects
            .select_related("imagem", "usuario", "novo_status")
            .order_by("-criado_em")[:8]
        )
        ctx.update({
            "visao": "coordenacao",
            "total": total,
            "finalizadas": finalizadas,
            "pendentes": pendentes,
            "por_status": por_status,
            "historico_recente": historico_recente,
        })

    elif usuario.tipo == usuario.Tipo.DESCRITOR:
        minhas = Imagem.objects.filter(responsavel=usuario, ativo=True).select_related("status")
        disponiveis = minhas.filter(status__slug__in=["liberado-descricao", "descrevendo"])
        ctx.update({
            "visao": "descritor",
            "total_minhas": minhas.count(),
            "disponiveis": disponiveis,
            "disponiveis_count": disponiveis.count(),
        })

    elif usuario.tipo == usuario.Tipo.REVISOR:
        minhas = Imagem.objects.filter(responsavel=usuario, ativo=True).select_related("status")
        para_conferir = minhas.filter(status__slug__in=["liberado-conferencia", "em-conferencia"])
        ctx.update({
            "visao": "revisor",
            "total_minhas": minhas.count(),
            "para_conferir": para_conferir,
            "para_conferir_count": para_conferir.count(),
        })

    return render(request, "core/dashboard.html", ctx)

# ============================================================
# MINHAS TAREFAS
# ============================================================

@login_required
def minhas_tarefas(request):
    from django.core.paginator import Paginator

    usuario = request.user
    status_slug = request.GET.get("status", "")
    busca = request.GET.get("busca", "").strip()
    pagina = request.GET.get("pagina", 1)

    # Definir quais status e título conforme perfil
    if usuario.tipo == usuario.Tipo.DESCRITOR:
        slugs_visiveis = ["liberado-descricao", "descrevendo", "descrito"]
        titulo_secao = "Minhas tarefas de descrição"
        acao_label = "Descrever"
        acao_icon = "bi-pencil-square"

    elif usuario.tipo == usuario.Tipo.REVISOR:
        slugs_visiveis = [
            "liberado-conferencia", "em-conferencia", "conferido"
        ]
        titulo_secao = "Minhas tarefas de conferência"
        acao_label = "Conferir"
        acao_icon = "bi-eye"

    elif usuario.tipo in (usuario.Tipo.COORDENADOR, usuario.Tipo.ADMINISTRADOR):
        slugs_visiveis = [
            "liberado-descricao", "descrevendo", "descrito",
            "liberado-conferencia", "em-conferencia", "conferido",
            "revisando", "revisado", "finalizado",
        ]
        titulo_secao = "Todas as tarefas"
        acao_label = "Abrir"
        acao_icon = "bi-arrow-right-circle"

    else:
        slugs_visiveis = []
        titulo_secao = "Minhas tarefas"
        acao_label = "Abrir"
        acao_icon = "bi-arrow-right-circle"

    # Queryset base
    if usuario.tipo in (usuario.Tipo.COORDENADOR, usuario.Tipo.ADMINISTRADOR):
        tarefas = Imagem.objects.filter(
            ativo=True,
            status__slug__in=slugs_visiveis,
        )
    else:
        tarefas = Imagem.objects.filter(
            ativo=True,
            responsavel=usuario,
            status__slug__in=slugs_visiveis,
        )

    # Filtros
    if status_slug:
        tarefas = tarefas.filter(status__slug=status_slug)
    if busca:
        tarefas = tarefas.filter(retranca__icontains=busca)

    tarefas = tarefas.select_related("status", "responsavel").order_by("-criado_em")

    # Contadores por status (para os cards de resumo)
    from django.db.models import Count, Q
    contadores = {}
    for slug in slugs_visiveis:
        if usuario.tipo in (usuario.Tipo.COORDENADOR, usuario.Tipo.ADMINISTRADOR):
            contadores[slug] = Imagem.objects.filter(
                ativo=True, status__slug=slug
            ).count()
        else:
            contadores[slug] = Imagem.objects.filter(
                ativo=True, responsavel=usuario, status__slug=slug
            ).count()

    total = tarefas.count()
    paginador = Paginator(tarefas, 50)
    pagina_obj = paginador.get_page(pagina)

    # Status disponíveis para filtro
    status_list = StatusWorkflow.objects.filter(
        ativo=True, slug__in=slugs_visiveis
    ).order_by("ordem")

    ctx = {
        "tarefas": pagina_obj,
        "pagina_obj": pagina_obj,
        "total": total,
        "titulo_secao": titulo_secao,
        "acao_label": acao_label,
        "acao_icon": acao_icon,
        "status_list": status_list,
        "status_slug_ativo": status_slug,
        "busca": busca,
        "contadores": contadores,
        "slugs_ativos": ["liberado-descricao", "descrevendo",
                         "liberado-conferencia", "em-conferencia,"],
        "mostrar_todos_status": usuario.tipo in (
            usuario.Tipo.COORDENADOR, usuario.Tipo.ADMINISTRADOR
        ),
    }
    return render(request, "core/minhas_tarefas.html", ctx)

# ============================================================
# IMAGENS — LISTAGEM
# ============================================================

@login_required
def imagens_lista(request):
    from django.core.paginator import Paginator

    status_slug = request.GET.get("status", "")
    busca = request.GET.get("busca", "").strip()
    pagina = request.GET.get("pagina", 1)

    imagens = Imagem.objects.filter(ativo=True).select_related("status", "responsavel")

    if status_slug:
        imagens = imagens.filter(status__slug=status_slug)
    if busca:
        imagens = imagens.filter(retranca__icontains=busca)

    imagens = imagens.order_by("-criado_em")

    total = imagens.count()
    paginador = Paginator(imagens, 50)  # 50 por página
    pagina_obj = paginador.get_page(pagina)

    status_list = StatusWorkflow.objects.filter(ativo=True).order_by("ordem")

    ctx = {
        "imagens": pagina_obj,
        "status_list": status_list,
        "status_slug_ativo": status_slug,
        "busca": busca,
        "total": total,
        "pagina_obj": pagina_obj,
    }
    return render(request, "core/imagens_lista.html", ctx)


# ============================================================
# IMAGENS — IMPORTAÇÃO VIA EXCEL
# ============================================================

@login_required
def importar_imagens(request):
    if not _apenas_coordenador(request.user):
        messages.error(request, "Apenas coordenadores e administradores podem importar imagens.")
        return redirect("dashboard")

    if request.method == "GET":
        return render(request, "core/importar_imagens.html")

    arquivo = request.FILES.get("arquivo")
    if not arquivo:
        messages.error(request, "Nenhum arquivo enviado.")
        return render(request, "core/importar_imagens.html")

    if not arquivo.name.endswith(".xlsx"):
        messages.error(request, "O arquivo deve estar no formato .xlsx")
        return render(request, "core/importar_imagens.html")

    try:
        status_inicial = StatusWorkflow.objects.get(slug="liberado-descricao")
    except StatusWorkflow.DoesNotExist:
        messages.error(request, "Status 'Liberado para descrição' não encontrado.")
        return render(request, "core/importar_imagens.html")

    try:
        wb = openpyxl.load_workbook(arquivo, read_only=True, data_only=True)
    except Exception as e:
        messages.error(request, f"Erro ao abrir o arquivo: {e}")
        return render(request, "core/importar_imagens.html")

    ws = wb.active
    headers = None
    data_rows = []

    for row in ws.iter_rows(values_only=True):
        if headers is None:
            headers = row
            continue
        data_rows.append(dict(zip(headers, row)))

    wb.close()

    if not data_rows:
        messages.error(request, "O arquivo está vazio.")
        return render(request, "core/importar_imagens.html")

    # ---- Pré-carregar dados de referência ----
    # Retrancas já existentes no banco
    retrancas_existentes = set(
        Imagem.objects.values_list("retranca", flat=True)
    )

    # Cache de usuários por username
    usernames = {
        str(r.get("usuario", "") or "").strip()
        for r in data_rows
        if r.get("usuario")
    }
    usuarios_cache = {
        u.username: u
        for u in Usuario.objects.filter(username__in=usernames)
    }

    # Cache de idiomas pycountry
    idioma_cache = {}

    def resolver_idioma(lang_code):
        if lang_code in idioma_cache:
            return idioma_cache[lang_code]
        alpha_2 = lang_code.split("-")[0]
        try:
            lang = pycountry.languages.get(alpha_2=alpha_2)
            codigo = lang.alpha_3 if lang and hasattr(lang, "alpha_3") else "und"
            nome = lang.name if lang else lang_code
        except Exception:
            codigo, nome = "und", lang_code
        idioma_cache[lang_code] = (codigo, nome)
        return codigo, nome

    # ---- Processar linhas ----
    imagens_criar = []
    rows_com_descricao = []
    puladas = 0
    erros = []
    log = []

    etapas_validas = [e[0] for e in Imagem.Etapa.choices]

    for data in data_rows:
        retranca = str(data.get("retranca") or "").strip()
        if not retranca:
            puladas += 1
            continue

        if retranca in retrancas_existentes:
            puladas += 1
            log.append({"tipo": "pulada", "retranca": retranca, "msg": "Já existe"})
            continue

        username = str(data.get("usuario") or "").strip()
        responsavel = usuarios_cache.get(username)

        etapa_raw = str(data.get("etapa") or "AD")
        etapa = etapa_raw.replace("Etapa:", "").strip()
        if etapa not in etapas_validas:
            etapa = Imagem.Etapa.AD

        img_file = str(data.get("img_file") or "")
        nome_arquivo = os.path.basename(img_file) if img_file else ""

        imagem = Imagem(
            retranca=retranca,
            nome_obra=str(data.get("colecao") or ""),
            componente_curricular=str(data.get("disciplina") or ""),
            volume_ano_modulo=str(data.get("volume") or ""),
            capitulo_unidade=str(data.get("capitulo") or ""),
            etapa=etapa,
            nome_arquivo=nome_arquivo,
            caminho_arquivo=img_file,
            status=status_inicial,
            responsavel=responsavel,
            cadastrado_por=request.user,
            ativo=True,
        )
        imagens_criar.append(imagem)
        rows_com_descricao.append((retranca, responsavel, data.get("descricao")))
        retrancas_existentes.add(retranca)  # evitar duplicatas no mesmo arquivo
        log.append({"tipo": "criada", "retranca": retranca, "msg": "Importada com sucesso"})

    # ---- Bulk create imagens em lotes de 500 ----
    LOTE = 500
    criadas = 0

    try:
        with transaction.atomic():
            for i in range(0, len(imagens_criar), LOTE):
                lote = imagens_criar[i:i + LOTE]
                Imagem.objects.bulk_create(lote, ignore_conflicts=True)
                criadas += len(lote)

            # Buscar IDs das imagens criadas
            retrancas_criadas = [img.retranca for img in imagens_criar]
            imagens_db = {
                img.retranca: img
                for img in Imagem.objects.filter(retranca__in=retrancas_criadas)
            }

            # ---- Bulk create descrições ----
            descricoes_criar = []
            trechos_por_retranca = {}

            for retranca, responsavel, descricao_raw in rows_com_descricao:
                imagem = imagens_db.get(retranca)
                if not imagem or not descricao_raw:
                    continue

                try:
                    trechos_data = ast.literal_eval(str(descricao_raw))
                    if not isinstance(trechos_data, list) or not trechos_data:
                        continue
                except (ValueError, SyntaxError):
                    continue

                descricao = Descricao(
                    imagem=imagem,
                    descritor=responsavel,
                    descritor_bloqueado=False,
                )
                descricoes_criar.append(descricao)
                trechos_por_retranca[retranca] = trechos_data

            for i in range(0, len(descricoes_criar), LOTE):
                Descricao.objects.bulk_create(descricoes_criar[i:i + LOTE], ignore_conflicts=True)

            # Buscar IDs das descrições criadas
            descricoes_db = {
                d.imagem.retranca: d
                for d in Descricao.objects.filter(
                    imagem__retranca__in=list(trechos_por_retranca.keys())
                ).select_related("imagem")
            }

            # ---- Bulk create trechos ----
            trechos_criar = []

            for retranca, trechos_data in trechos_por_retranca.items():
                descricao = descricoes_db.get(retranca)
                if not descricao:
                    continue

                for ordem, trecho_data in enumerate(trechos_data, 1):
                    lang_code = trecho_data.get("lang", "pt-BR")
                    texto = trecho_data.get("text", "")
                    idioma_codigo, idioma_nome = resolver_idioma(lang_code)

                    trechos_criar.append(Trecho(
                        descricao=descricao,
                        ordem=ordem,
                        texto=texto,
                        idioma_codigo=idioma_codigo,
                        idioma_nome=idioma_nome,
                    ))

            for i in range(0, len(trechos_criar), LOTE):
                Trecho.objects.bulk_create(trechos_criar[i:i + LOTE])

    except Exception as e:
        erros.append({"retranca": "—", "msg": str(e)})
        criadas = 0

    ctx = {
        "resultado": True,
        "criadas": criadas,
        "puladas": puladas,
        "erros": erros,
        "log": log[:200],  # limita o log a 200 linhas para não travar o navegador
        "log_truncado": len(log) > 200,
        "total_log": len(log),
    }
    return render(request, "core/importar_imagens.html", ctx)

# ============================================================
# DESCRIÇÃO DA IMAGEM
# ============================================================

@login_required
def descricao_imagem(request, pk):
    from django.shortcuts import get_object_or_404

    imagem = get_object_or_404(Imagem, pk=pk, ativo=True)
    usuario = request.user
    descricao = getattr(imagem, "descricao", None)

    # ---- Verificar permissão de acesso ----
    pode_editar = False
    motivo_bloqueio = None

    if usuario.tipo in (usuario.Tipo.ADMINISTRADOR, usuario.Tipo.COORDENADOR):
        pode_editar = True

    elif usuario.tipo == usuario.Tipo.DESCRITOR:
        if imagem.status.slug not in ("liberado-descricao", "descrevendo"):
            motivo_bloqueio = "Esta imagem não está disponível para descrição."
        elif descricao and descricao.descritor_bloqueado:
            motivo_bloqueio = "Seu acesso a esta descrição foi bloqueado após o salvamento. Somente o coordenador pode liberar novamente."
        elif imagem.responsavel and imagem.responsavel != usuario:
            motivo_bloqueio = "Esta tarefa está atribuída a outro descritor."
        else:
            pode_editar = True

    elif usuario.tipo == usuario.Tipo.REVISOR:
        if imagem.status.slug not in ("liberado-conferencia", "em-conferencia"):
            motivo_bloqueio = "Esta imagem não está disponível para conferência."
        else:
            pode_editar = True

    else:
        motivo_bloqueio = "Você não tem permissão para acessar esta tela."

    # ---- Buscar trechos existentes ----
    trechos = []
    if descricao:
        trechos = descricao.trechos.filter(ativo=True).order_by("ordem")

    # ---- Todos os idiomas via pycountry ----
    import json

    PRIORITARIOS = ["por", "eng", "spa", "fra", "deu", "ita", "jpn",
                    "zho", "lat", "ara", "rus", "hin", "kor", "grk"]

    todos = []
    prioritarios = []
    for lang in pycountry.languages:
        if not hasattr(lang, "alpha_3"):
            continue
        entry = {"codigo": lang.alpha_3, "nome": lang.name}
        if lang.alpha_3 in PRIORITARIOS:
            prioritarios.append(entry)
        else:
            todos.append(entry)

    prioritarios.sort(key=lambda x: x["nome"])
    todos.sort(key=lambda x: x["nome"])
    idiomas_json = json.dumps(prioritarios + todos, ensure_ascii=False)

    ctx = {
        "imagem": imagem,
        "descricao": descricao,
        "trechos": trechos,
        "pode_editar": pode_editar,
        "motivo_bloqueio": motivo_bloqueio,
        "idiomas_json": idiomas_json,
    }
    return render(request, "core/descricao.html", ctx)

import json
from django.http import JsonResponse
from django.views.decorators.http import require_POST


@login_required
@require_POST
@login_required
@require_POST
def salvar_trecho(request, pk):
    """Salva ou atualiza os trechos de uma descrição via AJAX."""
    from django.shortcuts import get_object_or_404
    from .models import HistoricoItem

    imagem = get_object_or_404(Imagem, pk=pk, ativo=True)
    usuario = request.user

    # Verificar permissão
    pode_editar = False
    if usuario.tipo in (usuario.Tipo.ADMINISTRADOR, usuario.Tipo.COORDENADOR):
        pode_editar = True
    elif usuario.tipo == usuario.Tipo.DESCRITOR:
        descricao_atual = getattr(imagem, "descricao", None)
        if (imagem.status.slug in ("liberado-descricao", "descrevendo")
                and (not descricao_atual or not descricao_atual.descritor_bloqueado)
                and (not imagem.responsavel or imagem.responsavel == usuario)):
            pode_editar = True
    elif usuario.tipo == usuario.Tipo.REVISOR:
        if imagem.status.slug in ("liberado-conferencia", "em-conferencia"):
            pode_editar = True

    if not pode_editar:
        return JsonResponse({"ok": False, "erro": "Sem permissão."}, status=403)

    try:
        body = json.loads(request.body)
    except json.JSONDecodeError:
        return JsonResponse({"ok": False, "erro": "JSON inválido."}, status=400)

    descricao_existia = hasattr(imagem, "descricao")

    # Garantir que a Descrição existe
    descricao, criada = Descricao.objects.get_or_create(
        imagem=imagem,
        defaults={"descritor": usuario if usuario.tipo == usuario.Tipo.DESCRITOR else None,
                  "descritor_bloqueado": False},
    )

    trechos_data = body.get("trechos", [])

    with transaction.atomic():
        descricao.trechos.all().delete()

        novos = []
        for i, t in enumerate(trechos_data, 1):
            texto = t.get("texto", "").strip()
            idioma_codigo = t.get("idioma_codigo", "por")
            idioma_nome = t.get("idioma_nome", "Português")
            if texto:
                novos.append(Trecho(
                    descricao=descricao,
                    ordem=i,
                    texto=texto,
                    idioma_codigo=idioma_codigo,
                    idioma_nome=idioma_nome,
                ))

        Trecho.objects.bulk_create(novos)

        # ---- Registrar histórico: primeira descrição criada pelo Descritor ----
        if usuario.tipo == usuario.Tipo.DESCRITOR and criada:
            HistoricoItem.objects.create(
                imagem=imagem,
                descricao=descricao,
                usuario=usuario,
                tipo_acao=HistoricoItem.TipoAcao.DESCRICAO_INICIADA,
                observacao="Descrição iniciada pelo descritor.",
            )

    return JsonResponse({"ok": True, "total": len(novos)})

@login_required
@require_POST
@login_required
@require_POST
def avancar_status(request, pk):
    """Avança o status da imagem no workflow, registrando o tipo_acao específico."""
    from django.shortcuts import get_object_or_404
    from .models import HistoricoItem

    imagem = get_object_or_404(Imagem, pk=pk, ativo=True)
    usuario = request.user

    # Mapa de transição: slug atual -> (próximo slug, tipo_acao do HistoricoItem)
    FLUXO = {
        "liberado-descricao":   ("descrevendo",           HistoricoItem.TipoAcao.DESCRICAO_INICIADA),
        "descrevendo":          ("descrito",              HistoricoItem.TipoAcao.DESCRICAO_SALVA),
        "descrito":             ("liberado-conferencia",  HistoricoItem.TipoAcao.LIBERADO_CONFERENCIA),
        "liberado-conferencia": ("em-conferencia",         HistoricoItem.TipoAcao.CONFERENCIA_INICIADA),
        "em-conferencia":       ("conferido",             HistoricoItem.TipoAcao.CONFERENCIA_CONCLUIDA),
        "conferido":            ("revisando",             HistoricoItem.TipoAcao.REVISAO_INICIADA),
        "revisando":            ("revisado",              HistoricoItem.TipoAcao.REVISAO_CONCLUIDA),
        "revisado":             ("finalizado",            HistoricoItem.TipoAcao.DESCRICAO_FINALIZADA),
    }

    transicao = FLUXO.get(imagem.status.slug)
    if not transicao:
        return JsonResponse({"ok": False, "erro": "Status final atingido."}, status=400)

    proximo_slug, tipo_acao = transicao

    try:
        proximo_status = StatusWorkflow.objects.get(slug=proximo_slug)
    except StatusWorkflow.DoesNotExist:
        return JsonResponse({"ok": False, "erro": "Status não encontrado."}, status=400)

    status_anterior = imagem.status
    descricao = getattr(imagem, "descricao", None)

    with transaction.atomic():
        imagem.status = proximo_status
        imagem.save()

        # ---- Atualizar responsáveis da Descrição conforme a etapa ----
        if descricao:
            if proximo_slug == "descrevendo":
                if not descricao.descritor:
                    descricao.descritor = usuario
                descricao.save()
            elif proximo_slug == "em-conferencia":
                if not descricao.revisor:
                    descricao.revisor = usuario
                descricao.save()
            elif proximo_slug == "revisando":
                if not descricao.coordenador:
                    descricao.coordenador = usuario
                descricao.save()
            elif proximo_slug == "finalizado":
                descricao.finalizado = True
                descricao.save()

        HistoricoItem.objects.create(
            imagem=imagem,
            descricao=descricao,
            usuario=usuario,
            tipo_acao=tipo_acao,
            status_anterior=status_anterior,
            novo_status=proximo_status,
        )

    return JsonResponse({
        "ok": True,
        "novo_status": proximo_status.nome,
        "novo_slug": proximo_status.slug,
    })
# ============================================================
# TESTE
# ============================================================

@login_required
def teste(request):
    return render(request, "core/teste.html")

# ============================================================
#  CRUD
# ============================================================

def _pode_gerenciar_imagem(usuario):
    """Regra de negócio: só Coordenador e Administrador gerenciam imagens."""
    return usuario.tipo in (Usuario.Tipo.ADMINISTRADOR, Usuario.Tipo.COORDENADOR)


@login_required
def imagem_criar(request):
    """Create — cadastra uma nova imagem individual."""
    if not _pode_gerenciar_imagem(request.user):
        messages.error(request, "Você não tem permissão para cadastrar imagens.")
        return redirect("imagens_lista")

    if request.method == "POST":
        form = ImagemForm(request.POST)
        if form.is_valid():
            imagem = form.save(commit=False)
            imagem.cadastrado_por = request.user
            imagem.save()
            messages.success(request, "Imagem cadastrada com sucesso.")
            return redirect("imagens_lista")
    else:
        # Pré-seleciona o primeiro status do workflow (Liberado para descrição)
        status_inicial = StatusWorkflow.objects.order_by("ordem").first()
        form = ImagemForm(initial={"status": status_inicial})

    return render(request, "core/imagem_form.html", {
        "form": form,
        "titulo": "Cadastrar imagem",
    })


@login_required
def imagem_editar(request, pk):
    """Update — edita os dados de uma imagem existente."""
    imagem = get_object_or_404(Imagem, pk=pk, ativo=True)

    if not _pode_gerenciar_imagem(request.user):
        messages.error(request, "Você não tem permissão para editar imagens.")
        return redirect("imagens_lista")

    if request.method == "POST":
        form = ImagemForm(request.POST, instance=imagem)
        if form.is_valid():
            form.save()
            messages.success(request, "Imagem atualizada com sucesso.")
            return redirect("imagens_lista")
    else:
        form = ImagemForm(instance=imagem)

    return render(request, "core/imagem_form.html", {
        "form": form,
        "titulo": "Editar imagem",
        "imagem": imagem,
    })


@login_required
def imagem_excluir(request, pk):
    """Delete — exclusão LÓGICA (marca ativo=False, não apaga do banco)."""
    imagem = get_object_or_404(Imagem, pk=pk, ativo=True)

    if not _pode_gerenciar_imagem(request.user):
        messages.error(request, "Você não tem permissão para excluir imagens.")
        return redirect("imagens_lista")

    if request.method == "POST":
        imagem.ativo = False
        imagem.save()
        messages.success(request, "Imagem excluída com sucesso.")
        return redirect("imagens_lista")

    return render(request, "core/imagem_excluir.html", {
        "imagem": imagem,
    })