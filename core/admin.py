from django.contrib import admin
from django.contrib.auth.admin import UserAdmin
from .models import Usuario


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