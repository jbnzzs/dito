"""Limite de segurança dos downloads XLSX do Dito!.

O número de linhas é SEMPRE calculado com os filtros do request, jamais
com o tamanho da página atualmente visível no navegador.
"""

from functools import wraps

from django.contrib import messages
from django.shortcuts import redirect
from django.urls import reverse

LIMITE_XLSX = 200


def limitar_xlsx_200(view):
    @wraps(view)
    def protegida(request, *args, **kwargs):
        # Import local evita referência circular com core.views.
        from .views import _imagens_relatorio

        qs = _imagens_relatorio(request)
        # Limitamos a contagem do banco a 201 registros (sem trazer linhas).
        quantidade = qs.order_by().values_list("pk", flat=True)[:LIMITE_XLSX + 1].count()
        if quantidade > LIMITE_XLSX:
            messages.warning(
                request,
                "O Excel aceita no máximo 200 imagens por exportação. "
                "Aplique mais filtros ou utilize o CSV para resultados maiores.",
            )
            destino = reverse("relatorios_lista")
            parametros = request.GET.urlencode()
            if parametros:
                destino = f"{destino}?{parametros}"
            return redirect(destino)
        return view(request, *args, **kwargs)

    return protegida
