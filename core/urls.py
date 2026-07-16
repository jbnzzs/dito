from django.urls import path
from . import views

urlpatterns = [
    path("", views.dashboard, name="dashboard"),
    path("minhas-tarefas/", views.minhas_tarefas, name="minhas_tarefas"),

    # ---- Imagens ----
    path("imagens/", views.imagens_lista, name="imagens_lista"),
    path("imagens/nova/", views.imagem_criar, name="imagem_criar"),
    path("imagens/importar/", views.importar_imagens, name="importar_imagens"),
    path("imagens/importar/<uuid:importacao_id>/lotes/", views.organizar_lotes, name="organizar_lotes"),
    path("imagens/<int:pk>/editar/", views.imagem_editar, name="imagem_editar"),
    path("imagens/<int:pk>/excluir/", views.imagem_excluir, name="imagem_excluir"),
    path("imagens/<int:pk>/descrever/", views.descricao_imagem, name="descricao_imagem"),
    path("imagens/<int:pk>/salvar-trecho/", views.salvar_trecho, name="salvar_trecho"),
    path("imagens/<int:pk>/avancar-status/", views.avancar_status, name="avancar_status"),
    path("imagens/<int:pk>/atribuir-descritor/", views.atribuir_descritor, name="atribuir_descritor"),
    path("imagens/<int:pk>/liberar-revisor/", views.liberar_conferencia, name="liberar_conferencia"),
    path("imagens/<int:pk>/atualizar-pagamento/", views.atualizar_pagamento, name="atualizar_pagamento"),
    path("imagens/<int:pk>/proxima-do-lote/", views.proxima_imagem_lote, name="proxima_imagem_lote"),
    path("imagens/<int:pk>/devolver-lote/", views.devolver_lote, name="devolver_lote"),

    # ---- Lotes ----
    path("lotes/", views.lotes_lista, name="lotes_lista"),
    path("lotes/organizar/", views.organizar_lotes, name="organizar_lotes_geral"),
    path("lotes/<int:lote_id>/atribuir/", views.atribuir_lote, name="atribuir_lote"),
    path("lotes/<int:pk>/editar/", views.lote_editar, name="lote_editar"),

    path("teste/", views.teste, name="teste"),
]