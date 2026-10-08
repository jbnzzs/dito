from django.contrib.auth.models import AbstractUser, BaseUserManager
from django.db import models


def filtro_autoria_imagem(usuario):
    """
    Q() que identifica imagens que "pertencem" ao usuário — seja porque ele
    é o responsável atual (tarefa em andamento) ou porque ele foi o autor
    daquela fase no passado (descricao.descritor/.revisor), mesmo depois
    do handoff para a próxima fase. Sem isso, tarefas e lotes "somem" da
    visão da pessoa assim que ela entrega o trabalho — que é o sintoma
    que este helper corrige.
    """
    from django.db.models import Q

    if usuario.tipo == Usuario.Tipo.DESCRITOR:
        return Q(responsavel=usuario) | Q(descricao__descritor=usuario)
    elif usuario.tipo == Usuario.Tipo.REVISOR:
        return Q(responsavel=usuario) | Q(descricao__revisor=usuario)
    return Q(responsavel=usuario)


# ============================================================
# USUÁRIO CUSTOMIZADO
# ============================================================

class UsuarioManager(BaseUserManager):
    """Manager que cria usuários usando e-mail como identificador principal."""
    use_in_migrations = True

    def _gerar_username(self, email):
        # username interno derivado do e-mail; o Django ainda o exige internamente
        base = email.split("@")[0]
        username = base
        contador = 1
        while self.model.objects.filter(username=username).exists():
            username = f"{base}{contador}"
            contador += 1
        return username

    def create_user(self, email, password=None, **extra_fields):
        if not email:
            raise ValueError("O e-mail é obrigatório.")
        email = self.normalize_email(email)
        extra_fields.setdefault("username", self._gerar_username(email))
        user = self.model(email=email, **extra_fields)
        user.set_password(password)
        user.save(using=self._db)
        return user

    def create_superuser(self, email, password=None, **extra_fields):
        extra_fields.setdefault("is_staff", True)
        extra_fields.setdefault("is_superuser", True)
        extra_fields.setdefault("tipo", "administrador")
        if extra_fields.get("is_staff") is not True:
            raise ValueError("Superusuário precisa de is_staff=True.")
        if extra_fields.get("is_superuser") is not True:
            raise ValueError("Superusuário precisa de is_superuser=True.")
        return self.create_user(email, password, **extra_fields)


class Usuario(AbstractUser):

    class Tipo(models.TextChoices):
        ADMINISTRADOR = "administrador", "Administrador"
        COORDENADOR = "coordenador", "Coordenador"
        DESCRITOR = "descritor", "Descritor"
        REVISOR = "revisor", "Revisor"

    class Situacao(models.TextChoices):
        PENDENTE = "pendente", "Pendente de aprovação"
        APROVADO = "aprovado", "Aprovado"
        RECUSADO = "recusado", "Recusado"

    email = models.EmailField("E-mail", unique=True)

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

    situacao = models.CharField(
        max_length=20,
        choices=Situacao.choices,
        default=Situacao.APROVADO,
        verbose_name="Situação do cadastro",
        help_text="Solicitações pela tela pública nascem como 'Pendente'.",
    )
    aprovado_por = models.ForeignKey(
        "self",
        on_delete=models.SET_NULL,
        null=True,
        blank=True,
        related_name="usuarios_aprovados",
        verbose_name="Aprovado/recusado por",
    )
    decidido_em = models.DateTimeField(
        null=True,
        blank=True,
        verbose_name="Data da decisão",
    )
    observacao_decisao = models.TextField(
        blank=True,
        verbose_name="Observação da decisão",
        help_text="Motivo da recusa ou observação do administrador.",
    )

    objects = UsuarioManager()

    USERNAME_FIELD = "email"
    REQUIRED_FIELDS = []

    class Meta:
        verbose_name = "Usuário"
        verbose_name_plural = "Usuários"

    def __str__(self):
        return f"{self.get_full_name() or self.email} ({self.get_tipo_display()})"

    @property
    def contrato_ativo(self):
        """Retorna True se o usuário possui contrato ativo hoje."""
        from datetime import date
        hoje = date.today()
        if self.contrato_inicio and self.contrato_fim:
            return self.contrato_inicio <= hoje <= self.contrato_fim
        return True

    @property
    def esta_pendente(self):
        return self.situacao == self.Situacao.PENDENTE

    @property
    def foi_recusado(self):
        return self.situacao == self.Situacao.RECUSADO

# ============================================================
# WORKFLOW
# ============================================================

class StatusWorkflow(models.Model):
    """
    Status do fluxo editorial. Gerenciável via Django Admin.
    Populado automaticamente pelo management command seed_status.
    """

    class PerfilResponsavel(models.TextChoices):
        DESCRITOR = "descritor", "Descritor"
        REVISOR = "revisor", "Revisor"
        COORDENADOR = "coordenador", "Coordenador"

    nome = models.CharField(max_length=60, unique=True, verbose_name="Nome")
    slug = models.SlugField(max_length=60, unique=True, verbose_name="Identificador interno")
    ordem = models.PositiveSmallIntegerField(verbose_name="Ordem de exibição")
    ativo = models.BooleanField(default=True, verbose_name="Ativo")
    descricao = models.TextField(blank=True, verbose_name="Descrição")

    perfil_responsavel = models.CharField(
        max_length=20,
        choices=PerfilResponsavel.choices,
        default=PerfilResponsavel.DESCRITOR,
        verbose_name="Perfil responsável",
        help_text="Qual perfil trabalha neste status.",
    )
    exige_atribuicao = models.BooleanField(
        default=False,
        verbose_name="Exige atribuição do coordenador",
        help_text="Se marcado, o coordenador precisa escolher um responsável antes de avançar.",
    )
    permite_edicao = models.BooleanField(
        default=False,
        verbose_name="Permite edição",
        help_text="Se marcado, o perfil responsável pode editar a descrição neste status.",
    )
    avanca_ao_abrir = models.BooleanField(
        default=False,
        verbose_name="Avança ao abrir",
        help_text="Se marcado, a tarefa avança automaticamente para o próximo status ao ser aberta pelo responsável.",
    )
    descricao_concluida = models.BooleanField(
        default=False,
        verbose_name="Descrição concluída",
        help_text="Identifica que este status representa a conclusão da etapa de descrição.",
    )
    conferencia_concluida = models.BooleanField(
        default=False,
        verbose_name="Conferência concluída",
        help_text="Identifica que este status representa a conclusão da etapa de conferência.",
    )
    revisao_concluida = models.BooleanField(
        default=False,
        verbose_name="Revisão concluída",
        help_text="Identifica que este status representa a conclusão da etapa de revisão.",
    )
    is_inicial = models.BooleanField(
        default=False,
        verbose_name="Status inicial",
        help_text="O status em que as imagens entram ao serem importadas.",
    )
    is_final = models.BooleanField(
        default=False,
        verbose_name="Status final",
        help_text="O status que encerra o fluxo (nada avança a partir dele).",
    )

    class Meta:
        verbose_name = "Status do workflow"
        verbose_name_plural = "Status do workflow"
        ordering = ["ordem"]

    def __str__(self):
        return f"{self.ordem}. {self.nome}"

    def proximo(self):
        """
        Retorna o próximo status ativo na fila (maior ordem que a atual),
        ou None se este for o último. Base do fluxo configurável.
        """
        return (
            StatusWorkflow.objects.filter(ativo=True, ordem__gt=self.ordem)
            .order_by("ordem")
            .first()
        )

    def anterior(self):
        """Retorna o status ativo imediatamente anterior na fila, ou None."""
        return (
            StatusWorkflow.objects.filter(ativo=True, ordem__lt=self.ordem)
            .order_by("-ordem")
            .first()
        )


# ============================================================
# CADASTROS EDITORIAIS
# ============================================================

class Projeto(models.Model):
    """
    Agrupador editorial de alto nível.

    Um mesmo projeto pode reunir imagens de diferentes obras/editoras.
    A inativação não apaga o registro nem rompe vínculos existentes.
    """
    nome = models.CharField(
        max_length=200,
        unique=True,
        verbose_name="Nome do projeto",
    )
    descricao = models.CharField(
        max_length=255,
        blank=True,
        verbose_name="Descrição",
    )
    acervo_fotoweb = models.CharField(
        max_length=180,
        blank=True,
        verbose_name="Acervo FotoWeb",
        help_text=(
            "Nome do acervo usado para gerar automaticamente "
            "os links das imagens no FotoWeb."
        ),
    )
    ativo = models.BooleanField(default=True, verbose_name="Ativo")
    criado_em = models.DateTimeField(auto_now_add=True, verbose_name="Criado em")
    atualizado_em = models.DateTimeField(auto_now=True, verbose_name="Atualizado em")

    class Meta:
        verbose_name = "Projeto"
        verbose_name_plural = "Projetos"
        ordering = ["nome"]

    def __str__(self):
        return self.nome


class ComponenteCurricular(models.Model):
    """Cadastro gerenciável dos componentes curriculares usados nas imagens."""

    nome = models.CharField(
        max_length=100,
        unique=True,
        verbose_name="Componente curricular",
    )
    ativo = models.BooleanField(default=True, verbose_name="Ativo")
    criado_em = models.DateTimeField(auto_now_add=True, verbose_name="Criado em")
    atualizado_em = models.DateTimeField(auto_now=True, verbose_name="Atualizado em")

    class Meta:
        verbose_name = "Componente curricular"
        verbose_name_plural = "Componentes curriculares"
        ordering = ["nome"]

    def __str__(self):
        return self.nome


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
        PENDENTE = "pendente", "Pendente"
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
    projeto = models.ForeignKey(
        Projeto,
        on_delete=models.PROTECT,
        null=True,
        blank=True,
        related_name="imagens",
        verbose_name="Projeto",
        help_text="Agrupador maior que pode reunir imagens de diferentes obras/editoras.",
    )
    nome_obra = models.CharField(max_length=200, verbose_name="Nome da obra")
    volume_ano_modulo = models.CharField(
        max_length=60,
        blank=True,
        verbose_name="Volume/Ano/Módulo",
    )
    componente_curricular = models.ForeignKey(
        ComponenteCurricular,
        on_delete=models.PROTECT,
        null=True,
        blank=True,
        related_name="imagens",
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
    dados_fotoweb_originais = models.JSONField(
        default=dict,
        blank=True,
        verbose_name="Dados originais do FotoWeb",
        help_text=(
            "Snapshot da linha original importada do relatório FotoWeb. "
            "Usado para gerar um novo arquivo no mesmo formato, alterando "
            "somente os campos de descrição."
        ),
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
        default=StatusPagamento.PENDENTE,
        verbose_name="Pagamento — Descritor",
    )
    pagamento_revisor = models.CharField(
        max_length=20,
        choices=StatusPagamento.choices,
        default=StatusPagamento.PENDENTE,
        verbose_name="Pagamento — Revisor",
    )

    # ---- Datas e controle ----
    prazo = models.DateField(null=True, blank=True, verbose_name="Prazo final")
    criado_em = models.DateTimeField(auto_now_add=True, verbose_name="Cadastrado em")
    atualizado_em = models.DateTimeField(auto_now=True, verbose_name="Atualizado em")
    ativo = models.BooleanField(default=True, verbose_name="Ativo")
    pronto_para_lote = models.BooleanField(
        default=False,
        verbose_name="Pronto para envio em lote",
        help_text="Marca que o usuário já concluiu esta imagem e está aguardando "
                   "o envio do lote inteiro, sem avançar o status individualmente.",
    )

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
    revisor_bloqueado = models.BooleanField(
        default=False,
        verbose_name="Revisor bloqueado",
        help_text="Bloqueado automaticamente após a conferência ser concluída.",
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
        REVISOR_BLOQUEADO        = "revisor_bloqueado",        "Acesso do revisor bloqueado"
        LIBERADO_CONFERENCIA     = "liberado_conferencia",     "Liberado para conferência"
        CONFERENCIA_INICIADA     = "conferencia_iniciada",     "Conferência iniciada"
        CONFERENCIA_CONCLUIDA    = "conferencia_concluida",    "Conferência concluída"
        DEVOLVIDO_CORRECAO       = "devolvido_correcao",       "Devolvido para correção"
        DESCRITOR_LIBERADO       = "descritor_liberado",       "Acesso do descritor liberado"
        REVISOR_LIBERADO         = "revisor_liberado",          "Acesso do revisor liberado"
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
        help_text="Ex: 'MAT V2'.",
    )
    descricao = models.CharField(
        max_length=255,
        blank=True,
        verbose_name="Descrição",
        help_text="Observação opcional sobre o critério usado para formar o lote.",
    )
    data_prevista = models.DateField(
        null=True,
        blank=True,
        verbose_name="Data prevista",
        help_text="Previsão de conclusão do lote.",
    )
    data_efetiva = models.DateField(
        null=True,
        blank=True,
        verbose_name="Data efetiva",
        help_text=(
            "Data em que todas as imagens ativas do lote chegaram ao status final. "
            "É preenchida automaticamente, mas pode ser ajustada pela coordenação."
        ),
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
        """
        True quando todas as imagens ativas do lote já alcançaram o marco
        configurado como conclusão da descrição. O nome e a slug do status
        não participam da regra; a posição do marco no workflow é usada para
        considerar também imagens que já avançaram para etapas posteriores.
        """
        imagens = self.imagens.filter(ativo=True)
        if not imagens.exists():
            return False

        marco_descricao = (
            StatusWorkflow.objects
            .filter(ativo=True, descricao_concluida=True)
            .order_by("ordem")
            .first()
        )
        if not marco_descricao:
            return False

        return not imagens.filter(status__ordem__lt=marco_descricao.ordem).exists()

    def progresso_do_usuario(self, usuario):
        """
        Progresso pessoal de um descritor/revisor dentro do lote. Considera
        tanto as imagens atualmente atribuídas quanto as que ele já entregou
        (autoria preservada em descricao.descritor/.revisor), para que o
        lote continue visível em modo consulta depois do envio.

        Uma imagem é considerada pendente para o perfil quando o status atual
        pertence a esse perfil e representa uma etapa operacional: permite
        edição ou avança automaticamente ao ser aberta. Assim, o cálculo não
        depende mais de nomes ou slugs específicos do workflow.
        """
        from django.db.models import Q

        if usuario.tipo not in (Usuario.Tipo.DESCRITOR, Usuario.Tipo.REVISOR):
            return None

        minhas = (
            self.imagens
            .filter(ativo=True)
            .filter(filtro_autoria_imagem(usuario))
            .distinct()
        )
        total = minhas.count()

        if total == 0:
            return None

        restantes = (
            minhas
            .filter(status__perfil_responsavel=usuario.tipo)
            .filter(
                Q(status__permite_edicao=True)
                | Q(status__avanca_ao_abrir=True)
            )
            .count()
        )
        concluidas = total - restantes

        return {
            "total": total,
            "concluidas": concluidas,
            "restantes": restantes,
            "percentual": round((concluidas / total) * 100, 1) if total else 0,
        }

    class Meta:
        verbose_name = "Lote"
        verbose_name_plural = "Lotes"
        ordering = ["-criado_em"]

    def __str__(self):
        return self.nome

    @property
    def total_imagens(self):
        return self.imagens.filter(ativo=True).count()

    def esta_totalmente_finalizado(self):
        """
        True quando o lote possui imagens ativas e todas estão em um status
        marcado com is_final=True. O nome do status não participa da regra.
        """
        imagens = self.imagens.filter(ativo=True)
        if not imagens.exists():
            return False
        return not imagens.exclude(status__is_final=True).exists()

    def sincronizar_data_efetiva(self):
        """
        Preenche a data efetiva na primeira vez em que todas as imagens ativas
        do lote chegam ao status final. Uma data já registrada não é sobrescrita.
        """
        if self.data_efetiva or not self.esta_totalmente_finalizado():
            return False

        from django.utils import timezone

        self.data_efetiva = timezone.localdate()
        self.save(update_fields=["data_efetiva"])
        return True

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
            .values_list("status_id")
            .annotate(qtd=Count("id"))
        )

        resultado = []
        for status in StatusWorkflow.objects.filter(ativo=True).order_by("ordem"):
            qtd = contagens.get(status.pk, 0)
            if qtd:
                resultado.append({
                    "status_id": status.pk,
                    "slug": status.slug,
                    "nome": status.nome,
                    "ordem": status.ordem,
                    "quantidade": qtd,
                    "percentual": round((qtd / total) * 100, 1),
                })
        return resultado

# ============================================================
# SINCRONIZAÇÃO AUTOMÁTICA DE DATAS DO LOTE
# ============================================================

from django.db.models.signals import post_save
from django.dispatch import receiver


@receiver(post_save, sender=Imagem)
def sincronizar_data_efetiva_lote_ao_salvar_imagem(sender, instance, **kwargs):
    """
    Sempre que uma imagem vinculada a um lote é salva, verifica se o lote
    acabou de ser concluído. A data efetiva é gravada apenas uma vez.
    """
    if not instance.lote_id:
        return

    instance.lote.sincronizar_data_efetiva()

