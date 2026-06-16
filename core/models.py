from django.contrib.auth.models import AbstractUser
from django.db import models


# ============================================================
# USUÁRIO CUSTOMIZADO
# ============================================================

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
        return True


# ============================================================
# WORKFLOW
# ============================================================

class StatusWorkflow(models.Model):
    """
    Status do fluxo editorial. Gerenciável via Django Admin.
    Populado automaticamente pelo management command seed_status.
    """
    nome = models.CharField(max_length=60, unique=True, verbose_name="Nome")
    slug = models.SlugField(max_length=60, unique=True, verbose_name="Identificador interno")
    ordem = models.PositiveSmallIntegerField(verbose_name="Ordem de exibição")
    ativo = models.BooleanField(default=True, verbose_name="Ativo")
    descricao = models.TextField(blank=True, verbose_name="Descrição")

    class Meta:
        verbose_name = "Status do workflow"
        verbose_name_plural = "Status do workflow"
        ordering = ["ordem"]

    def __str__(self):
        return f"{self.ordem}. {self.nome}"


# ============================================================
# IMAGEM
# ============================================================

class Imagem(models.Model):
    """
    Imagem cadastrada no sistema. Centro do fluxo editorial.
    """

    class Etapa(models.TextChoices):
        AD       = "AD",       "AD"
        FT       = "FT",       "FT — Fechamento Técnico"
        ORIGINAL = "Original", "Original"
        P1       = "P1",       "P1 — Prova 1"
        P2       = "P2",       "P2 — Prova 2"
        P3       = "P3",       "P3 — Prova 3"

    # ---- Metadados editoriais ----
    retranca = models.CharField(
        max_length=120,
        verbose_name="Retranca",
        help_text="Identificador editorial da imagem.",
    )
    nome_obra = models.CharField(max_length=200, verbose_name="Nome da obra")
    volume_ano_modulo = models.CharField(
        max_length=60,
        blank=True,
        verbose_name="Volume / Ano / Módulo",
    )
    componente_curricular = models.CharField(
        max_length=100,
        blank=True,
        verbose_name="Componente curricular",
    )
    capitulo_unidade = models.CharField(
        max_length=100,
        blank=True,
        verbose_name="Capítulo / Unidade",
    )
    etapa = models.CharField(
        max_length=10,
        choices=Etapa.choices,
        default=Etapa.AD,
        verbose_name="Etapa",
    )

    # ---- Informações técnicas ----
    nome_arquivo = models.CharField(
        max_length=255,
        blank=True,
        verbose_name="Nome do arquivo",
    )
    tamanho_arquivo = models.CharField(
        max_length=30,
        blank=True,
        verbose_name="Tamanho do arquivo",
    )
    dimensoes = models.CharField(
        max_length=30,
        blank=True,
        verbose_name="Dimensões",
        help_text="Ex: 1920x1080 px",
    )
    resolucao = models.CharField(
        max_length=20,
        blank=True,
        verbose_name="Resolução",
        help_text="Ex: 300 dpi",
    )
    tamanho_fisico = models.CharField(
        max_length=30,
        blank=True,
        verbose_name="Tamanho físico",
        help_text="Ex: 15x10 cm",
    )

    # ---- Workflow ----
    status = models.ForeignKey(
        StatusWorkflow,
        on_delete=models.PROTECT,
        verbose_name="Status atual",
        related_name="imagens",
    )
    responsavel = models.ForeignKey(
        Usuario,
        on_delete=models.SET_NULL,
        null=True,
        blank=True,
        related_name="imagens_responsavel",
        verbose_name="Responsável",
    )
    cadastrado_por = models.ForeignKey(
        Usuario,
        on_delete=models.SET_NULL,
        null=True,
        related_name="imagens_cadastradas",
        verbose_name="Cadastrado por",
    )

    # ---- Datas e controle ----
    prazo = models.DateField(null=True, blank=True, verbose_name="Prazo final")
    criado_em = models.DateTimeField(auto_now_add=True, verbose_name="Cadastrado em")
    atualizado_em = models.DateTimeField(auto_now=True, verbose_name="Atualizado em")
    ativo = models.BooleanField(default=True, verbose_name="Ativo")

    class Meta:
        verbose_name = "Imagem"
        verbose_name_plural = "Imagens"
        ordering = ["-criado_em"]

    def __str__(self):
        return f"{self.retranca} — {self.nome_obra}"


# ============================================================
# DESCRIÇÃO
# ============================================================

class Descricao(models.Model):
    """
    Descrição vinculada a uma imagem.
    Uma imagem possui uma descrição.
    """
    imagem = models.OneToOneField(
        Imagem,
        on_delete=models.CASCADE,
        related_name="descricao",
        verbose_name="Imagem",
    )

    # ---- Responsáveis por etapa ----
    descritor = models.ForeignKey(
        Usuario,
        on_delete=models.SET_NULL,
        null=True,
        blank=True,
        related_name="descricoes_produzidas",
        verbose_name="Descritor",
    )
    revisor = models.ForeignKey(
        Usuario,
        on_delete=models.SET_NULL,
        null=True,
        blank=True,
        related_name="descricoes_revisadas",
        verbose_name="Revisor",
    )
    coordenador = models.ForeignKey(
        Usuario,
        on_delete=models.SET_NULL,
        null=True,
        blank=True,
        related_name="descricoes_coordenadas",
        verbose_name="Coordenador",
    )

    # ---- Datas por etapa ----
    inicio_descricao = models.DateTimeField(null=True, blank=True, verbose_name="Início da descrição")
    fim_descricao = models.DateTimeField(null=True, blank=True, verbose_name="Fim da descrição")
    inicio_conferencia = models.DateTimeField(null=True, blank=True, verbose_name="Início da conferência")
    fim_conferencia = models.DateTimeField(null=True, blank=True, verbose_name="Fim da conferência")
    inicio_revisao = models.DateTimeField(null=True, blank=True, verbose_name="Início da revisão final")
    fim_revisao = models.DateTimeField(null=True, blank=True, verbose_name="Fim da revisão final")
    finalizado_em = models.DateTimeField(null=True, blank=True, verbose_name="Finalizado em")

    # ---- Controle de acesso ----
    descritor_bloqueado = models.BooleanField(
        default=False,
        verbose_name="Descritor bloqueado",
        help_text="Bloqueado automaticamente após o primeiro salvamento.",
    )

    # ---- Conteúdo e observações ----
    observacoes = models.TextField(blank=True, verbose_name="Observações internas")
    finalizado = models.BooleanField(default=False, verbose_name="Finalizado")

    # ---- Datas de controle ----
    criado_em = models.DateTimeField(auto_now_add=True)
    atualizado_em = models.DateTimeField(auto_now=True)

    class Meta:
        verbose_name = "Descrição"
        verbose_name_plural = "Descrições"

    def __str__(self):
        return f"Descrição de {self.imagem.retranca}"


# ============================================================
# TRECHO
# ============================================================

class Trecho(models.Model):
    """
    Parte da descrição com idioma identificado.
    Uma descrição pode ter vários trechos ordenados.
    """
    descricao = models.ForeignKey(
        Descricao,
        on_delete=models.CASCADE,
        related_name="trechos",
        verbose_name="Descrição",
    )
    ordem = models.PositiveSmallIntegerField(verbose_name="Ordem")
    texto = models.TextField(verbose_name="Texto do trecho")

    # Idioma via pycountry (código ISO 639-3)
    idioma_codigo = models.CharField(
        max_length=10,
        default="por",
        verbose_name="Código do idioma",
        help_text="Código ISO 639-3. Ex: por (Português), eng (Inglês).",
    )
    idioma_nome = models.CharField(
        max_length=100,
        default="Português",
        verbose_name="Nome do idioma",
    )

    ativo = models.BooleanField(default=True, verbose_name="Ativo")
    criado_em = models.DateTimeField(auto_now_add=True)
    atualizado_em = models.DateTimeField(auto_now=True)

    class Meta:
        verbose_name = "Trecho"
        verbose_name_plural = "Trechos"
        ordering = ["ordem"]

    def __str__(self):
        return f"Trecho {self.ordem} — {self.idioma_nome} ({self.descricao.imagem.retranca})"


# ============================================================
# HISTÓRICO
# ============================================================

class HistoricoItem(models.Model):
    """
    Registro de ações relevantes no fluxo editorial.
    Garante rastreabilidade e auditoria.
    """

    class TipoAcao(models.TextChoices):
        IMAGEM_CADASTRADA        = "imagem_cadastrada",        "Imagem cadastrada"
        TAREFA_ATRIBUIDA         = "tarefa_atribuida",         "Tarefa atribuída"
        DESCRICAO_INICIADA       = "descricao_iniciada",       "Descrição iniciada"
        DESCRICAO_SALVA          = "descricao_salva",          "Descrição salva"
        DESCRITOR_BLOQUEADO      = "descritor_bloqueado",      "Descritor bloqueado"
        LIBERADO_CONFERENCIA     = "liberado_conferencia",     "Liberado para conferência"
        CONFERENCIA_INICIADA     = "conferencia_iniciada",     "Conferência iniciada"
        CONFERENCIA_CONCLUIDA    = "conferencia_concluida",    "Conferência concluída"
        DEVOLVIDO_CORRECAO       = "devolvido_correcao",       "Devolvido para correção"
        DESCRITOR_LIBERADO       = "descritor_liberado",       "Acesso do descritor liberado"
        REVISAO_INICIADA         = "revisao_iniciada",         "Revisão final iniciada"
        REVISAO_CONCLUIDA        = "revisao_concluida",        "Revisão final concluída"
        DESCRICAO_FINALIZADA     = "descricao_finalizada",     "Descrição finalizada"
        ENVIADO_FOTOWEB          = "enviado_fotoweb",          "Enviado ao FotoWeb"
        STATUS_ALTERADO          = "status_alterado",          "Status alterado"

    imagem = models.ForeignKey(
        Imagem,
        on_delete=models.CASCADE,
        related_name="historico",
        verbose_name="Imagem",
    )
    descricao = models.ForeignKey(
        Descricao,
        on_delete=models.SET_NULL,
        null=True,
        blank=True,
        related_name="historico",
        verbose_name="Descrição",
    )
    usuario = models.ForeignKey(
        Usuario,
        on_delete=models.SET_NULL,
        null=True,
        related_name="acoes",
        verbose_name="Usuário",
    )
    tipo_acao = models.CharField(
        max_length=40,
        choices=TipoAcao.choices,
        verbose_name="Tipo de ação",
    )
    status_anterior = models.ForeignKey(
        StatusWorkflow,
        on_delete=models.SET_NULL,
        null=True,
        blank=True,
        related_name="historico_anterior",
        verbose_name="Status anterior",
    )
    novo_status = models.ForeignKey(
        StatusWorkflow,
        on_delete=models.SET_NULL,
        null=True,
        blank=True,
        related_name="historico_novo",
        verbose_name="Novo status",
    )
    observacao = models.TextField(
        blank=True,
        verbose_name="Observação",
        help_text="Justificativa ou detalhe da ação.",
    )
    criado_em = models.DateTimeField(auto_now_add=True, verbose_name="Data e hora")

    class Meta:
        verbose_name = "Item de histórico"
        verbose_name_plural = "Histórico"
        ordering = ["-criado_em"]

    def __str__(self):
        return f"{self.get_tipo_acao_display()} — {self.imagem.retranca} ({self.criado_em:%d/%m/%Y %H:%M})"