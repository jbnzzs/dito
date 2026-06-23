from django.urls import path
from . import views

urlpatterns = [
    path("", views.dashboard, name="dashboard"),
    path("minhas-tarefas/", views.minhas_tarefas, name="minhas_tarefas"),
    path("imagens/", views.imagens_lista, name="imagens_lista"),
    path("imagens/importar/", views.importar_imagens, name="importar_imagens"),
    path("imagens/<int:pk>/descrever/", views.descricao_imagem, name="descricao_imagem"),
    path("imagens/<int:pk>/salvar-trecho/", views.salvar_trecho, name="salvar_trecho"),
    path("imagens/<int:pk>/avancar-status/", views.avancar_status, name="avancar_status"),
    path("teste/", views.teste, name="teste"),
]