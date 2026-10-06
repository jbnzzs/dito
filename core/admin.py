from django.contrib import admin
from django.contrib.auth.admin import UserAdmin
from .models import (
    ComponenteCurricular,
    Descricao,
    HistoricoItem,
    Imagem,
    Lote,
    Projeto,
    StatusWorkflow,
    Trecho,
    Usuario,
)


@admin.register(Usuario)
class UsuarioAdmin(UserAdmin):
    # Login é por e-mail: o Admin precisa refletir isso
    ordering = ("-date_joined",)
    list_display = ("email", "get_full_name", "tipo", "situacao", "is_active", "date_joined")
    list_filter = ("tipo", "situacao", "is_active", "is_staff")
    search_fields = ("email", "first_name", "last_name")
    list_editable = ("tipo", "is_active")

    # Edição de um usuário existente
    fieldsets = (
        ("Acesso", {"fields": ("email", "password")}),
        ("Dados pessoais", {"fields": ("first_name", "last_name")}),
        ("Perfil no Dito!", {"fields": ("tipo", "situacao", "contrato_inicio", "contrato_fim")}),
        ("Aprovação", {"fields": ("aprovado_por", "decidido_em", "observacao_decisao")}),
        ("Permissões", {"fields": ("is_active", "is_staff", "is_superuser", "groups", "user_permissions")}),
        ("Datas", {"fields": ("last_login", "date_joined")}),
    )

    # Criação de um novo usuário (o Admin pede e-mail + senha, não username)
    add_fieldsets = (
        (None, {
            "classes": ("wide",),
            "fields": ("email", "password1", "password2", "tipo"),
        }),
    )

    readonly_fields = ("date_joined", "last_login", "decidido_em")

    actions = ["ativar_usuarios", "desativar_usuarios"]

    @admin.action(description="Ativar usuários selecionados")
    def ativar_usuarios(self, request, queryset):
        n = queryset.update(is_active=True)
        self.message_user(request, f"{n} usuário(s) ativado(s).")

    @admin.action(description="Desativar usuários selecionados")
    def desativar_usuarios(self, request, queryset):
        n = queryset.update(is_active=False)
        self.message_user(request, f"{n} usuário(s) desativado(s).")


@admin.register(StatusWorkflow)
class StatusWorkflowAdmin(admin.ModelAdmin):
    list_display = ("ordem", "nome", "slug", "ativo")
    list_editable = ("ativo",)
    ordering = ("ordem",)


@admin.register(Projeto)
class ProjetoAdmin(admin.ModelAdmin):
    list_display = ("nome", "ativo", "criado_em", "atualizado_em")
    list_filter = ("ativo",)
    search_fields = ("nome", "descricao")
    ordering = ("nome",)


@admin.register(ComponenteCurricular)
class ComponenteCurricularAdmin(admin.ModelAdmin):
    list_display = ("nome", "ativo", "criado_em", "atualizado_em")
    list_filter = ("ativo",)
    search_fields = ("nome",)
    ordering = ("nome",)


@admin.register(Imagem)
class ImagemAdmin(admin.ModelAdmin):
    list_display = ("retranca", "projeto", "nome_obra", "componente_curricular", "etapa", "status", "responsavel", "criado_em")
    list_filter = ("projeto", "componente_curricular", "status", "etapa", "ativo")
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