from django.contrib.auth.models import AbstractUser
from django.db import models


class Usuario(AbstractUser):

    class Tipo(models.TextChoices):
        ADMINISTRADOR = "administrador", "Administrador"
        COORDENADOR = "coordenador", "Coordenador"
        DESCRITOR = "descritor", "Descritor"
        REVISOR = "revisor", "Revisor"

    tipo = models.CharField(
        max_length=20,
        choices=Tipo.choices,
        default=Tipo.DESCRITOR,
        verbose_name="Tipo de perfil",
    )
    contrato_inicio = models.DateField(
        null=True,
        blank=True,
        verbose_name="Início do contrato",
    )
    contrato_fim = models.DateField(
        null=True,
        blank=True,
        verbose_name="Fim do contrato",
    )

    class Meta:
        verbose_name = "Usuário"
        verbose_name_plural = "Usuários"

    def __str__(self):
        return f"{self.get_full_name() or self.username} ({self.get_tipo_display()})"

    @property
    def contrato_ativo(self):
        """Retorna True se o usuário possui contrato ativo hoje."""
        from datetime import date
        hoje = date.today()
        if self.contrato_inicio and self.contrato_fim:
            return self.contrato_inicio <= hoje <= self.contrato_fim
        return True  # Administrador e Coordenador não têm contrato limitado