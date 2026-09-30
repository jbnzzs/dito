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
    path("imagens/<int:pk>/devolver-descritor/", views.devolver_descritor, name="devolver_descritor"),
    path("imagens/<int:pk>/devolver-revisor/", views.devolver_revisor, name="devolver_revisor"),
    path("solicitar-acesso/", views.solicitar_acesso, name="solicitar_acesso"),
    path("minha-conta/", views.minha_conta, name="minha_conta"),

    # ---- Lotes ----
    path("lotes/", views.lotes_lista, name="lotes_lista"),
    path("lotes/organizar/", views.organizar_lotes, name="organizar_lotes_geral"),
    path("lotes/<int:lote_id>/atribuir/", views.atribuir_lote, name="atribuir_lote"),
    path("lotes/<int:lote_id>/alterar-status/", views.lote_alterar_status, name="lote_alterar_status"),
    path("lotes/<int:pk>/editar/", views.lote_editar, name="lote_editar"),

    # ---- Usuários ----
    path("usuarios/", views.usuarios_lista, name="usuarios_lista"),
    path("usuarios/<int:pk>/aprovar/", views.aprovar_solicitacao, name="aprovar_solicitacao"),
    path("usuarios/<int:pk>/recusar/", views.recusar_solicitacao, name="recusar_solicitacao"),

    # ---- Status do workflow ----
    path("status/", views.status_lista, name="status_lista"),
    path("status/novo/", views.status_criar, name="status_criar"),
    path("status/<int:pk>/editar/", views.status_editar, name="status_editar"),
    path("status/<int:pk>/toggle-ativo/", views.status_toggle_ativo, name="status_toggle_ativo"),
    path("status/reordenar/", views.status_reordenar, name="status_reordenar"),

    # ---- Relatórios ----
    path("relatorios/", views.relatorios_lista, name="relatorios_lista"),
    path("relatorios/exportar/", views.relatorios_exportar, name="relatorios_exportar"),

    # ---- Buscar retranca ----
    path("buscar/", views.buscar_retranca, name="buscar_retranca"),

    # ---- Histórico ----
    path("historico/", views.historico_lista, name="historico_lista"),


]