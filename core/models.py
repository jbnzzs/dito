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

    class StatusPagamento(models.TextChoices):
        CONTABILIZADO = "contabilizado", "Contabilizado"
        PAGO = "pago", "Pago"

    # ---- Metadados editoriais ----
    retranca = models.CharField(
        max_length=120,
        unique=True,
        db_index=True,
        verbose_name="Retranca",
        help_text="Identificador editorial da imagem.",
    )
    nome_obra = models.CharField(max_length=200, verbose_name="Nome da obra")
    volume_ano_modulo = models.CharField(
        max_length=60,
        blank=True,
        verbose_name="Volume/Ano/Módulo",
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
    caminho_arquivo = models.CharField(
        max_length=500,
        blank=True,
        verbose_name="Caminho do arquivo",
        help_text="Caminho ou referência de onde o arquivo está armazenado.",
    )
    url_fotoweb = models.CharField(
        max_length=500,
        blank=True,
        verbose_name="URL no FotoWeb",
        help_text="Link para a página da imagem no FotoWeb (preview do asset).",
    )
    url_pdf = models.CharField(
        max_length=500,
        blank=True,
        verbose_name="URL do PDF",
        help_text="Link para o PDF da imagem.",
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

    # ---- Lote ----
    lote = models.ForeignKey(
        "Lote",
        on_delete=models.SET_NULL,
        null=True,
        blank=True,
        related_name="imagens",
        verbose_name="Lote",
    )
    importacao_id = models.UUIDField(
        null=True,
        blank=True,
        editable=False,
        verbose_name="ID da importação",
        help_text="Identifica todas as imagens vindas do mesmo upload de .xlsx, "
                   "permitindo isolar o lote de origem na tela de organização em lotes.",
    )

    # ---- Pagamento ----
    pagamento_descritor = models.CharField(
        max_length=20,
        choices=StatusPagamento.choices,
        default=StatusPagamento.CONTABILIZADO,
        verbose_name="Pagamento — Descritor",
    )
    pagamento_revisor = models.CharField(
        max_length=20,
        choices=StatusPagamento.choices,
        default=StatusPagamento.CONTABILIZADO,
        verbose_name="Pagamento — Revisor",
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
    imagem = models.OneToOneField(
        Imagem,
        on_delete=models.CASCADE,
        related_name="descricao",
        verbose_name="Imagem",
    )

    descritor = models.ForeignKey(
        Usuario, on_delete=models.SET_NULL, null=True, blank=True,
        related_name="descricoes_produzidas", verbose_name="Descritor",
    )
    revisor = models.ForeignKey(
        Usuario, on_delete=models.SET_NULL, null=True, blank=True,
        related_name="descricoes_revisadas", verbose_name="Revisor",
    )
    coordenador = models.ForeignKey(
        Usuario, on_delete=models.SET_NULL, null=True, blank=True,
        related_name="descricoes_coordenadas", verbose_name="Coordenador",
    )

    descritor_bloqueado = models.BooleanField(
        default=False,
        verbose_name="Descritor bloqueado",
        help_text="Bloqueado automaticamente após o primeiro salvamento.",
    )

    observacoes = models.TextField(blank=True, verbose_name="Observações internas")
    finalizado = models.BooleanField(default=False, verbose_name="Finalizado")

    criado_em = models.DateTimeField(auto_now_add=True)
    atualizado_em = models.DateTimeField(auto_now=True)

    class Meta:
        verbose_name = "Descrição"
        verbose_name_plural = "Descrições"

    def __str__(self):
        return f"Descrição de {self.imagem.retranca}"

    # ---- Datas por etapa (derivadas do HistoricoItem, não armazenadas) ----
    def _data_evento(self, tipo_acao, ultimo=False):
        qs = self.historico.filter(tipo_acao=tipo_acao).order_by(
            "-criado_em" if ultimo else "criado_em"
        )
        item = qs.first()
        return item.criado_em if item else None

    @property
    def inicio_descricao(self):
        return self._data_evento("descricao_iniciada")

    @property
    def fim_descricao(self):
        return self._data_evento("descricao_salva", ultimo=True)

    @property
    def inicio_conferencia(self):
        return self._data_evento("conferencia_iniciada")

    @property
    def fim_conferencia(self):
        return self._data_evento("conferencia_concluida")

    @property
    def inicio_revisao(self):
        return self._data_evento("revisao_iniciada")

    @property
    def fim_revisao(self):
        return self._data_evento("revisao_concluida")

    @property
    def finalizado_em(self):
        return self._data_evento("descricao_finalizada")

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
        unique_together = [("descricao", "ordem")]

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
        PAGAMENTO_ATUALIZADO     = "pagamento_atualizado",      "Status de pagamento atualizado"

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

# ============================================================
# LOTES
# ============================================================

class Lote(models.Model):
    """
    Agrupamento manual de imagens, criado logo após a importação de um
    arquivo .xlsx, para facilitar a atribuição de tarefas em bloco
    (em vez de imagem por imagem) para descritores e revisores.
    """
    nome = models.CharField(
        max_length=100,
        unique=True,
        verbose_name="Nome do lote",
        help_text="Ex: 'Matemática V2 2026'.",
    )
    descricao = models.CharField(
        max_length=255,
        blank=True,
        verbose_name="Descrição",
        help_text="Observação opcional sobre o critério usado para formar o lote.",
    )
    prazo = models.DateField(
        null=True,
        blank=True,
        verbose_name="Prazo do lote",
        help_text="Ao salvar, este prazo é aplicado a todas as imagens do lote.",
    )
    criado_por = models.ForeignKey(
        Usuario,
        on_delete=models.PROTECT,
        related_name="lotes_criados",
        verbose_name="Criado por",
    )
    criado_em = models.DateTimeField(auto_now_add=True, verbose_name="Criado em")
    ativo = models.BooleanField(
        default=True,
        verbose_name="Ativo",
        help_text="Exclusão lógica: lotes inativos não devem ser exibidos nas listagens.",
    )

    def esta_totalmente_descrito(self):
        """True se todas as imagens ativas do lote estão com status 'descrito'."""
        imagens = self.imagens.filter(ativo=True)
        if not imagens.exists():
            return False
        return not imagens.exclude(status__slug="descrito").exists()

    def liberar_para_conferencia(self, usuario):
        """
        Avança TODAS as imagens ativas do lote de 'descrito' para
        'liberado-conferencia' de uma vez, registrando histórico por imagem.
        Retorna a quantidade de imagens liberadas.
        """
        from django.db import transaction

        proximo = StatusWorkflow.objects.get(slug="liberado-conferencia")
        imagens = list(self.imagens.filter(ativo=True, status__slug="descrito"))

        if not imagens:
            return 0

        with transaction.atomic():
            for imagem in imagens:
                status_anterior = imagem.status
                imagem.status = proximo
                imagem.responsavel = None
                imagem.save()
                HistoricoItem.objects.create(
                    imagem=imagem,
                    descricao=getattr(imagem, "descricao", None),
                    usuario=usuario,
                    tipo_acao=HistoricoItem.TipoAcao.LIBERADO_CONFERENCIA,
                    status_anterior=status_anterior,
                    novo_status=proximo,
                    observacao=f"Liberado automaticamente: lote '{self.nome}' totalmente descrito.",
                )
        return len(imagens)

    class Meta:
        verbose_name = "Lote"
        verbose_name_plural = "Lotes"
        ordering = ["-criado_em"]

    def __str__(self):
        return self.nome

    @property
    def total_imagens(self):
        return self.imagens.filter(ativo=True).count()

    def propagar_prazo(self):
        """
        Aplica o prazo do lote a todas as imagens ativas vinculadas a ele.
        Retorna a quantidade de imagens atualizadas.
        """
        if self.prazo is None:
            return 0
        return self.imagens.filter(ativo=True).update(prazo=self.prazo)

    def progresso_por_status(self):
        """
        Retorna a distribuição das imagens ativas do lote pelos status do
        workflow, na ordem oficial, com contagem e percentual de cada um.
        Usado para montar a barra de progresso na tela de Lotes.
        """
        from django.db.models import Count

        total = self.total_imagens
        if total == 0:
            return []

        contagens = dict(
            self.imagens.filter(ativo=True)
            .values_list("status__slug")
            .annotate(qtd=Count("id"))
        )

        resultado = []
        for status in StatusWorkflow.objects.filter(ativo=True).order_by("ordem"):
            qtd = contagens.get(status.slug, 0)
            if qtd:
                resultado.append({
                    "slug": status.slug,
                    "nome": status.nome,
                    "ordem": status.ordem,
                    "quantidade": qtd,
                    "percentual": round((qtd / total) * 100, 1),
                })
        return resultado