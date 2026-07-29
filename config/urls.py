from django.contrib import admin
from django.urls import path, include

urlpatterns = [
    path("admin/", admin.site.urls),
    path("", include("django.contrib.auth.urls")),
    path("", include("core.urls")),
]

admin.site.site_header = "Dito! — Administração"
admin.site.site_title = "Dito! Admin"
admin.site.index_title = "Painel administrativo"