from django.shortcuts import render
from django.contrib.auth.decorators import login_required
from django.db.models import Count, Q
from datetime import datetime
from .models import Imagem, StatusWorkflow, HistoricoItem


def _saudacao():
    hora = datetime.now().hour
    if hora < 12:
        return "Bom dia"
    elif hora < 18:
        return "Boa tarde"
    return "Boa noite"


@login_required
def dashboard(request):
    usuario = request.user
    hoje = datetime.now().strftime("%d/%m/%Y")
    ctx = {
        "saudacao": _saudacao(),
        "hoje": hoje,
    }

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
        minhas = Imagem.objects.filter(
            responsavel=usuario, ativo=True
        ).select_related("status")
        disponiveis = minhas.filter(
            status__slug__in=["liberado-descricao", "descrevendo"]
        )
        ctx.update({
            "visao": "descritor",
            "total_minhas": minhas.count(),
            "disponiveis": disponiveis,
            "disponiveis_count": disponiveis.count(),
        })

    elif usuario.tipo == usuario.Tipo.REVISOR:
        minhas = Imagem.objects.filter(
            responsavel=usuario, ativo=True
        ).select_related("status")
        para_conferir = minhas.filter(
            status__slug__in=["liberado-conferencia", "em-conferencia"]
        )
        ctx.update({
            "visao": "revisor",
            "total_minhas": minhas.count(),
            "para_conferir": para_conferir,
            "para_conferir_count": para_conferir.count(),
        })

    return render(request, "core/dashboard.html", ctx)


@login_required
def teste(request):
    return render(request, "core/teste.html")