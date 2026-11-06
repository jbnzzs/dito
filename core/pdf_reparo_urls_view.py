"""Ferramenta administrativa de reparo de links historicos no banco do Render."""
import csv
from io import StringIO

from django.contrib.auth.decorators import login_required
from django.db import transaction
from django.http import HttpResponse, HttpResponseForbidden, JsonResponse
from django.shortcuts import render
from django.views.decorators.http import require_http_methods

from .models import Imagem, Usuario
from .services.pdf_urls_legacy_repair import (
    corrigir_url_pdf_antiga,
    listar_candidatos,
)

_TAMANHO_LOTE = 200


def _csv_backup(candidatos):
    saida = StringIO()
    writer = csv.writer(saida, delimiter=";")
    writer.writerow(["imagem_id", "retranca", "url_pdf_original", "url_pdf_corrigida"])
    for linha in candidatos:
        writer.writerow(linha)
    resposta = HttpResponse(saida.getvalue(), content_type="text/csv; charset=utf-8")
    resposta["Content-Disposition"] = 'attachment; filename="backup_pdfs_e2_antes_correcao.csv"'
    resposta["Cache-Control"] = "no-store"
    return resposta


@login_required
@require_http_methods(["GET", "POST"])
def manutencao_urls_pdf(request):
    # Restrito ao administrador; nao disponibilizar para descritores, revisores
    # ou coordenadores. O endpoint POST continua protegido pelo CSRF do Django.
    if not (request.user.is_superuser or request.user.tipo == Usuario.Tipo.ADMINISTRADOR):
        return HttpResponseForbidden("Acesso exclusivo de administradores.")

    if request.method == "GET":
        candidatos = listar_candidatos()
        if request.GET.get("baixar_backup") == "1":
            return _csv_backup(candidatos)
        return render(request, "core/pdf_reparo_urls.html", {
            "total_candidatos": len(candidatos),
            "amostras": candidatos[:20],
        })

    if (request.POST.get("confirmacao") != "CORRIGIR PDF"
            or request.POST.get("backup_confirmado") != "sim"):
        return JsonResponse({"erro": "Confirme o backup e digite CORRIGIR PDF."}, status=400)

    # Apenas 200 por requisicao. O JavaScript chama novamente ate completar,
    # evitando timeouts do Render. A operacao e repetivel com seguranca.
    candidatos = listar_candidatos()[:_TAMANHO_LOTE]
    if not candidatos:
        return JsonResponse({"alteradas": 0, "restantes": 0})

    urls_por_pk = {pk: (antiga, nova) for pk, _, antiga, nova in candidatos}
    alteradas = []
    with transaction.atomic():
        imagens = list(
            Imagem.objects.select_for_update()
            .filter(pk__in=list(urls_por_pk), ativo=True)
            .only("id", "url_pdf")
        )
        for imagem in imagens:
            antiga, nova = urls_por_pk[imagem.pk]
            # Revalidar no banco sob bloqueio antes de substituir.
            if imagem.url_pdf != antiga:
                continue
            if corrigir_url_pdf_antiga(imagem.url_pdf) != nova:
                continue
            imagem.url_pdf = nova
            alteradas.append(imagem)
        if alteradas:
            Imagem.objects.bulk_update(alteradas, ["url_pdf"], batch_size=100)

    return JsonResponse({
        "alteradas": len(alteradas),
        "restantes": len(listar_candidatos()),
    })
