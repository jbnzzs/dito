import ast
import os

import openpyxl
import pycountry
from django.contrib import messages
from django.contrib.auth.decorators import login_required
from django.db import transaction
from django.shortcuts import redirect, render

from .models import Descricao, Imagem, StatusWorkflow, Trecho, Usuario


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
# TESTE
# ============================================================

@login_required
def teste(request):
    return render(request, "core/teste.html")