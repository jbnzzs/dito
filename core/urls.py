from django.urls import path
from . import views

urlpatterns = [
    path("", views.dashboard, name="dashboard"),
    path("imagens/", views.imagens_lista, name="imagens_lista"),
    path("imagens/importar/", views.importar_imagens, name="importar_imagens"),
    path("teste/", views.teste, name="teste"),
]