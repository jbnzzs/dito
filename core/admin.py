from django.contrib import admin
from django.contrib.auth.admin import UserAdmin
from .models import Usuario, StatusWorkflow, Imagem, Descricao, Trecho, HistoricoItem, Lote


@admin.register(Usuario)
class UsuarioAdmin(UserAdmin):
    list_display = ("username", "get_full_name", "email", "tipo", "is_active")
    list_filter = ("tipo", "is_active", "is_staff")
    search_fields = ("username", "first_name", "last_name", "email")
    fieldsets = UserAdmin.fieldsets + (
        ("Perfil Dito!", {
            "fields": ("tipo", "contrato_inicio", "contrato_fim"),
        }),
    )
    add_fieldsets = UserAdmin.add_fieldsets + (
        ("Perfil Dito!", {
            "fields": ("tipo", "contrato_inicio", "contrato_fim"),
        }),
    )


@admin.register(StatusWorkflow)
class StatusWorkflowAdmin(admin.ModelAdmin):
    list_display = ("ordem", "nome", "slug", "ativo")
    list_editable = ("ativo",)
    ordering = ("ordem",)


@admin.register(Imagem)
class ImagemAdmin(admin.ModelAdmin):
    list_display = ("retranca", "nome_obra", "etapa", "status", "responsavel", "criado_em")
    list_filter = ("status", "etapa", "ativo")
    search_fields = ("retranca", "nome_obra")
    ordering = ("-criado_em",)


@admin.register(Descricao)
class DescricaoAdmin(admin.ModelAdmin):
    list_display = ("imagem", "descritor", "revisor", "coordenador", "finalizado")
    list_filter = ("finalizado", "descritor_bloqueado")
    search_fields = ("imagem__retranca",)


@admin.register(Trecho)
class TrechoAdmin(admin.ModelAdmin):
    list_display = ("descricao", "ordem", "idioma_nome", "ativo")
    list_filter = ("idioma_nome", "ativo")
    ordering = ("descricao", "ordem")


@admin.register(HistoricoItem)
class HistoricoItemAdmin(admin.ModelAdmin):
    list_display = ("imagem", "tipo_acao", "usuario", "criado_em")
    list_filter = ("tipo_acao",)
    search_fields = ("imagem__retranca",)
    ordering = ("-criado_em",)
    readonly_fields = ("criado_em",)


@admin.register(Lote)
class LoteAdmin(admin.ModelAdmin):
    list_display = ("nome", "criado_por", "criado_em", "total_imagens", "ativo")
    list_filter = ("ativo",)
    search_fields = ("nome", "descricao")
    ordering = ("-criado_em",)
    readonly_fields = ("criado_por", "criado_em")