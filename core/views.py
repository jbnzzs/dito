import ast
import os
import re
import unicodedata
import uuid
from functools import lru_cache
from urllib.parse import quote, urlencode, urlsplit

import openpyxl
import pycountry
from babel import Locale
from django.contrib import messages
from django.contrib.auth.decorators import login_required
from django.db import transaction
from django.shortcuts import redirect, render, get_object_or_404
from django.urls import reverse
from django.utils.http import url_has_allowed_host_and_scheme

from .models import (
    ComponenteCurricular,
    Descricao,
    Imagem,
    Projeto,
    StatusWorkflow,
    Trecho,
    Usuario,
)
from .forms import ImagemForm


# ============================================================
# FOTOWEB — MODELO ORIGINAL DO RELATÓRIO
# ============================================================

FOTOWEB_COLUNAS_RELATORIO = [
    "obra",
    "componente",
    "volume",
    "capitulo",
    "keywords",
    "status",
    "retranca",
    "img_file",
    "descricao",
    "usuario",
    "etapa",
    "retranca_lower",
    "descricao_flat",
]


def _valor_json_fotoweb(valor):
    """
    Converte valores vindos do Excel para tipos seguros no JSONField.

    Strings, números, booleanos e nulos são mantidos. Datas/horas e
    outros objetos são convertidos para texto para que o snapshot possa
    ser persistido sem alterar os demais dados da linha.
    """
    if valor is None:
        return None

    if isinstance(
        valor,
        (
            str,
            int,
            float,
            bool,
        ),
    ):
        return valor

    if hasattr(valor, "isoformat"):
        try:
            return valor.isoformat()
        except Exception:
            pass

    return str(valor)


def _snapshot_linha_fotoweb(
    data,
    numero_linha=None,
    arquivo_origem="",
):
    """
    Guarda a linha original do relatório FotoWeb.

    Os campos internos iniciados por ``__`` são metadados do Dito e NÃO
    são exportados para o arquivo final.
    """
    snapshot = {
        coluna: _valor_json_fotoweb(
            data.get(coluna)
        )
        for coluna in FOTOWEB_COLUNAS_RELATORIO
    }

    snapshot["__numero_linha"] = numero_linha
    snapshot["__arquivo_origem"] = str(
        arquivo_origem
        or ""
    )

    return snapshot


# ============================================================
# EXCEL — NORMALIZAÇÃO DE COLUNAS
# ============================================================

def _normalizar_cabecalho_excel(valor):
    """
    Normaliza cabeçalhos vindos do Excel.

    Exemplos:
    - "Coleção" -> "colecao"
    - "Componente Curricular" -> "componente_curricular"
    - "IMG File" -> "img_file"

    Isso evita que diferenças de maiúsculas, acentos, espaços ou hífens
    façam os dados deixarem de ser importados.
    """
    texto = str(valor or "").strip()

    texto = unicodedata.normalize(
        "NFKD",
        texto,
    ).encode(
        "ascii",
        "ignore",
    ).decode(
        "ascii",
    )

    texto = texto.casefold()
    texto = re.sub(r"[^a-z0-9]+", "_", texto)

    return texto.strip("_")


def _valor_excel(data, *chaves):
    """
    Retorna o primeiro valor preenchido entre os nomes de coluna aceitos.
    """
    for chave in chaves:
        valor = data.get(chave)

        if valor is not None and str(valor).strip():
            return valor

    return ""


# ============================================================
# PDF — SHAREPOINT
# ============================================================

SHAREPOINT_SITE_URL = "https://ensinolivre.sharepoint.com/sites/G25Digital"
SHAREPOINT_BIBLIOTECA = "Shared Documents"
SHAREPOINT_PASTA_PDFS = "PDFs"
PDF_TIPO_MATERIAL = "mp"

PDF_COMPONENTES_SHAREPOINT = {
    "ARTE": "ART",
    "ART": "ART",
    "HISTORIA": "HIS",
    "HIS": "HIS",
}


def _normalizar_codigo_sharepoint(valor):
    texto = str(valor or "").strip()

    texto = unicodedata.normalize(
        "NFKD",
        texto,
    ).encode(
        "ascii",
        "ignore",
    ).decode(
        "ascii",
    )

    return re.sub(
        r"[^A-Za-z0-9]+",
        "",
        texto,
    ).upper()


def _codigo_componente_pdf_sharepoint(valor):
    normalizado = _normalizar_codigo_sharepoint(
        valor
    )

    if not normalizado:
        return ""

    return PDF_COMPONENTES_SHAREPOINT.get(
        normalizado,
        normalizado,
    )


def _normalizar_projeto_editorial_sharepoint(valor):
    partes = [
        _normalizar_codigo_sharepoint(parte)
        for parte in str(valor or "").strip().split("_")
        if str(parte or "").strip()
    ]

    return "_".join(
        parte
        for parte in partes
        if parte
    )


def _normalizar_numero_projeto_pdf(valor):
    numeros = re.sub(
        r"\D",
        "",
        str(valor or ""),
    )

    if not numeros:
        return ""

    return f"p{numeros.zfill(3)}"


def _extrair_ano_pdf(valor):
    encontrado = re.search(
        r"\d{1,2}",
        str(valor or ""),
    )

    if not encontrado:
        return ""

    return encontrado.group(0)


def _prefixo_arquivo_pdf(projeto_editorial):
    projeto = _normalizar_projeto_editorial_sharepoint(
        projeto_editorial
    )

    partes = [
        parte
        for parte in projeto.split("_")
        if parte
    ]

    if len(partes) < 2:
        return ""

    return "_".join(
        partes[:2]
    ).lower()


def _montar_url_pdf_sharepoint(
    projeto_editorial,
    numero_projeto,
    componente,
    volume_ano,
):
    projeto_editorial = (
        _normalizar_projeto_editorial_sharepoint(
            projeto_editorial
        )
    )

    partes = [
        parte
        for parte in projeto_editorial.split("_")
        if parte
    ]

    if len(partes) < 2:
        return ""

    projeto = partes[0].lower()
    prefixo = "_".join(partes[:2]).lower()

    numero = _normalizar_numero_projeto_pdf(
        numero_projeto
    )

    componente_codigo = (
        _codigo_componente_pdf_sharepoint(
            componente
        )
    )

    ano = _extrair_ano_pdf(
        volume_ano
    )

    if (
        not projeto
        or not prefixo
        or not numero
        or not componente_codigo
        or not ano
    ):
        return ""

    inicial_componente = (
        componente_codigo[0].lower()
    )

    nome_pdf = (
        f"{prefixo}_"
        f"{ano}"
        f"{numero}"
        f"{inicial_componente}_"
        f"{PDF_TIPO_MATERIAL}.pdf"
    )

    segmentos = [
        SHAREPOINT_BIBLIOTECA,
        projeto,
        SHAREPOINT_PASTA_PDFS,
        componente_codigo,
        nome_pdf,
    ]

    caminho = "/".join(
        quote(segmento, safe="")
        for segmento in segmentos
    )

    return (
        f"{SHAREPOINT_SITE_URL}/"
        f"{caminho}"
        "?web=1"
    )


# ============================================================
# FOTOWEB
# ============================================================

FOTOWEB_ARCHIVES_BASE_URL = (
    "http://fotoweb.ensinolivre.com.br:9090/fotoweb/archives/"
)


def _montar_url_fotoweb(acervo, retranca):
    """
    Monta a URL de pesquisa da imagem no FotoWeb.

    Cada relatório importado pertence a um único Acervo. A retranca da
    imagem é enviada no parâmetro ``q`` da pesquisa do FotoWeb.
    """
    acervo = str(acervo or "").strip().strip("/")
    retranca = str(retranca or "").strip()

    if not acervo or not retranca:
        return ""

    acervo_url = quote(acervo, safe="")
    retranca_url = quote(retranca, safe="")

    return (
        f"{FOTOWEB_ARCHIVES_BASE_URL}"
        f"{acervo_url}/?q={retranca_url}"
    )


@lru_cache(maxsize=1)
def _catalogo_idiomas_pt():
    """
    Retorna os idiomas que possuem nome localizado em português (pt-BR).

    O código ISO 639-3 continua sendo a referência interna. A lista exibida
    ao usuário vem do CLDR, por meio do Babel, para evitar que nomes sem
    tradução apareçam em inglês no seletor.
    """
    locale_pt = Locale.parse("pt_BR")
    nomes_por_codigo = {}

    for codigo_locale, nome_localizado in locale_pt.languages.items():
        idioma = None

        if len(codigo_locale) == 2:
            idioma = pycountry.languages.get(alpha_2=codigo_locale)
        elif len(codigo_locale) == 3:
            idioma = pycountry.languages.get(alpha_3=codigo_locale)

        if not idioma or not hasattr(idioma, "alpha_3"):
            continue

        nome = str(nome_localizado or "").strip()
        if not nome:
            continue

        # Na interface, mantemos inicial maiúscula para acompanhar o padrão
        # visual atual do seletor.
        nome = nome[:1].upper() + nome[1:]
        nomes_por_codigo.setdefault(idioma.alpha_3, nome)

    idiomas = [
        {"codigo": codigo, "nome": nome}
        for codigo, nome in nomes_por_codigo.items()
    ]
    idiomas.sort(key=lambda item: item["nome"].casefold())

    return idiomas, nomes_por_codigo


def _nome_idioma_pt(codigo, fallback=None):
    """Nome do idioma em português; nunca usa o nome inglês como fallback."""
    _, nomes_por_codigo = _catalogo_idiomas_pt()

    if codigo in nomes_por_codigo:
        return nomes_por_codigo[codigo]

    if codigo:
        return f"Idioma ({codigo.upper()})"

    return fallback or "Idioma não identificado"


def _saudacao():
    from datetime import datetime
    hora = datetime.now().hour
    if hora < 12:
        return "Bom dia"
    elif hora < 18:
        return "Boa tarde"
    return "Boa noite"

def _imagens_do_lote_para_usuario(lote, usuario):
    """
    Escopo de imagens pendentes de um lote para um usuário, conforme as
    propriedades configuradas no StatusWorkflow.

    A regra não depende do nome nem do slug do status:
    - Coordenador/Administrador: imagens não finalizadas.
    - Descritor/Revisor: imagens atribuídas a ele em status da sua fase que
      permitam edição ou avancem automaticamente ao abrir.
    """
    from django.db.models import Q

    if not lote:
        return Imagem.objects.none()

    base = Imagem.objects.filter(
        lote=lote,
        ativo=True,
        status__ativo=True,
    )

    if usuario.tipo in (Usuario.Tipo.COORDENADOR, Usuario.Tipo.ADMINISTRADOR):
        return (
            base
            .filter(status__is_final=False)
            .order_by("retranca")
        )

    if usuario.tipo in (Usuario.Tipo.DESCRITOR, Usuario.Tipo.REVISOR):
        return (
            base
            .filter(
                responsavel=usuario,
                status__perfil_responsavel=usuario.tipo,
            )
            .filter(
                Q(status__permite_edicao=True)
                | Q(status__avanca_ao_abrir=True)
            )
            .order_by("retranca")
        )

    return Imagem.objects.none()


def _proxima_imagem_do_lote(imagem_atual, usuario):
    """
    Próxima imagem pendente do mesmo lote para o perfil atual. Primeiro
    tenta a próxima na ordem (retranca maior); se não houver mais à frente,
    volta para o começo da fila — permite começar pelo meio do lote sem
    deixar imagens anteriores de fora.
    """
    if not imagem_atual.lote:
        return None

    escopo = _imagens_do_lote_para_usuario(imagem_atual.lote, usuario).exclude(pk=imagem_atual.pk)

    proxima = escopo.filter(retranca__gt=imagem_atual.retranca).order_by("retranca").first()
    if proxima:
        return proxima

    return escopo.order_by("retranca").first()

def _escopo_fixo_do_lote(lote, usuario):
    """
    Todas as imagens do lote que pertencem à fase deste usuário,
    independente do status atual.

    Esse escopo é usado como base estável para o total do lote na tela
    de descrição. A numeração exibida não representa mais a posição da
    retranca dentro do lote; ela representa quantas imagens o usuário
    já iniciou naquela etapa.
    """
    from django.db.models import Q

    if usuario.tipo in (Usuario.Tipo.COORDENADOR, Usuario.Tipo.ADMINISTRADOR):
        return Imagem.objects.filter(lote=lote, ativo=True).order_by("retranca")

    if usuario.tipo == Usuario.Tipo.DESCRITOR:
        return (
            Imagem.objects.filter(lote=lote, ativo=True)
            .filter(Q(responsavel=usuario) | Q(descricao__descritor=usuario))
            .distinct()
            .order_by("retranca")
        )

    if usuario.tipo == Usuario.Tipo.REVISOR:
        return (
            Imagem.objects.filter(lote=lote, ativo=True)
            .filter(Q(responsavel=usuario) | Q(descricao__revisor=usuario))
            .distinct()
            .order_by("retranca")
        )

    return Imagem.objects.none()

def _apenas_coordenador(usuario):
    return usuario.tipo in (usuario.Tipo.ADMINISTRADOR, usuario.Tipo.COORDENADOR)


def _url_interna_segura(request, url):
    """Retorna uma URL de retorno somente quando ela pertence ao próprio Dito!."""
    if not url:
        return None

    if not url_has_allowed_host_and_scheme(
        url,
        allowed_hosts={request.get_host()},
        require_https=request.is_secure(),
    ):
        return None

    return url


def _url_anterior_segura(request):
    """Retorna o Referer interno quando ele aponta para outra tela."""
    referer = _url_interna_segura(
        request,
        request.META.get("HTTP_REFERER"),
    )

    if not referer:
        return None

    try:
        path_referer = urlsplit(referer).path.rstrip("/")
        path_atual = request.path.rstrip("/")
    except (TypeError, ValueError):
        return None

    if path_referer == path_atual:
        return None

    return referer


def _perfil_operacional(usuario):
    """
    Perfil usado pelo motor do workflow.

    Administrador mantém acesso amplo ao sistema, mas quando participa de
    transições do fluxo atua como Coordenador.
    """
    if usuario.tipo == Usuario.Tipo.ADMINISTRADOR:
        return Usuario.Tipo.COORDENADOR
    return usuario.tipo


def _status_inicial_workflow():
    """Retorna o status inicial ativo configurado no workflow."""
    return (
        StatusWorkflow.objects
        .filter(ativo=True, is_inicial=True)
        .order_by("ordem")
        .first()
    )


def _status_entrada_perfil(perfil, depois_de=None):
    """
    Retorna o próximo status de entrada de um perfil.

    Um status de entrada é identificado pela flag avanca_ao_abrir, e não
    pelo nome/slug. Quando depois_de é informado, busca apenas status
    posteriores na fila.
    """
    qs = StatusWorkflow.objects.filter(
        ativo=True,
        perfil_responsavel=perfil,
        avanca_ao_abrir=True,
    )

    if depois_de is not None:
        qs = qs.filter(ordem__gt=depois_de.ordem)

    return qs.order_by("ordem").first()


def _usuario_pode_visualizar_imagem(usuario, imagem, descricao=None):
    """
    Coordenação/Admin podem visualizar qualquer imagem.
    Descritor/Revisor podem visualizar tarefas atualmente atribuídas a eles
    ou tarefas em que a autoria da etapa já foi registrada.
    """
    if usuario.tipo in (Usuario.Tipo.ADMINISTRADOR, Usuario.Tipo.COORDENADOR):
        return True

    if imagem.responsavel_id == usuario.id:
        return True

    if not descricao:
        return False

    if usuario.tipo == Usuario.Tipo.DESCRITOR:
        return descricao.descritor_id == usuario.id

    if usuario.tipo == Usuario.Tipo.REVISOR:
        return descricao.revisor_id == usuario.id

    return False


def _usuario_pode_iniciar_status(usuario, imagem, descricao=None):
    """Verifica se o usuário pode disparar a transição automática ao abrir."""
    status = imagem.status

    if not status or not status.ativo or not status.avanca_ao_abrir:
        return False

    if _perfil_operacional(usuario) != status.perfil_responsavel:
        return False

    if usuario.tipo == Usuario.Tipo.DESCRITOR:
        if imagem.responsavel_id != usuario.id:
            return False
        if descricao and descricao.descritor_bloqueado:
            return False

    if usuario.tipo == Usuario.Tipo.REVISOR:
        if imagem.responsavel_id != usuario.id:
            return False
        if descricao and descricao.revisor_bloqueado:
            return False

    return True


def _usuario_pode_editar_imagem(usuario, imagem, descricao=None):
    """
    A edição depende da flag permite_edicao do status.

    O nome e o slug do status não participam da regra de permissão.
    """
    status = imagem.status

    if not status or not status.ativo or not status.permite_edicao:
        return False

    # Administrador pode editar qualquer etapa explicitamente editável.
    if usuario.tipo == Usuario.Tipo.ADMINISTRADOR:
        return True

    if status.perfil_responsavel != usuario.tipo:
        return False

    if usuario.tipo == Usuario.Tipo.DESCRITOR:
        if imagem.responsavel_id != usuario.id:
            return False
        if descricao and descricao.descritor_bloqueado:
            return False

    elif usuario.tipo == Usuario.Tipo.REVISOR:
        if imagem.responsavel_id != usuario.id:
            return False
        if descricao and descricao.revisor_bloqueado:
            return False

    return True


def _tipo_acao_inicio(perfil):
    """Tipo de histórico gerado quando uma etapa começa ao abrir a tarefa."""
    from .models import HistoricoItem

    return {
        Usuario.Tipo.DESCRITOR: HistoricoItem.TipoAcao.DESCRICAO_INICIADA,
        Usuario.Tipo.REVISOR: HistoricoItem.TipoAcao.CONFERENCIA_INICIADA,
        Usuario.Tipo.COORDENADOR: HistoricoItem.TipoAcao.REVISAO_INICIADA,
    }.get(perfil, HistoricoItem.TipoAcao.STATUS_ALTERADO)


def _tipo_acao_transicao(status_anterior, novo_status):
    """
    Resolve a ação de histórico pelas flags semânticas do status de destino.
    """
    from .models import HistoricoItem

    if novo_status.descricao_concluida:
        return HistoricoItem.TipoAcao.DESCRICAO_SALVA

    if novo_status.conferencia_concluida:
        return HistoricoItem.TipoAcao.CONFERENCIA_CONCLUIDA

    if novo_status.revisao_concluida:
        return HistoricoItem.TipoAcao.REVISAO_CONCLUIDA

    if novo_status.is_final:
        return HistoricoItem.TipoAcao.DESCRICAO_FINALIZADA

    if status_anterior.avanca_ao_abrir and novo_status.permite_edicao:
        return _tipo_acao_inicio(novo_status.perfil_responsavel)

    return HistoricoItem.TipoAcao.STATUS_ALTERADO


def _normalizar_texto_descricao(texto):
    """Substitui aspas duplas por aspas simples antes de persistir."""
    return str(texto or "").replace('"', "'")


def _data_iso_ou_none(valor):
    """Converte YYYY-MM-DD vindo de input type=date para date ou None."""
    from datetime import date

    valor = str(valor or "").strip()
    if not valor:
        return None
    return date.fromisoformat(valor)


# ============================================================
# DASHBOARD
# ============================================================

@login_required
def dashboard(request):
    from datetime import datetime
    from django.db.models import Count, Q

    usuario = request.user
    hoje = datetime.now().strftime("%d/%m/%Y")
    ctx = {"saudacao": _saudacao(), "hoje": hoje}

    if usuario.tipo in (usuario.Tipo.ADMINISTRADOR, usuario.Tipo.COORDENADOR):
        total = Imagem.objects.filter(ativo=True).count()
        finalizadas = Imagem.objects.filter(
            ativo=True,
            status__is_final=True,
        ).count()
        pendentes = total - finalizadas

        por_status = (
            StatusWorkflow.objects
            .filter(ativo=True)
            .annotate(total=Count("imagens", filter=Q(imagens__ativo=True)))
            .order_by("ordem")
        )

        from .models import HistoricoItem

        historico_recente = (
            HistoricoItem.objects
            .select_related("imagem", "usuario", "novo_status")
            .order_by("-criado_em")[:8]
        )

        ctx.update({
            "visao": "coordenacao",
            "total": total,
            "finalizadas": finalizadas,
            "pendentes": pendentes,
            "por_status": por_status,
            "historico_recente": historico_recente,
        })

    elif usuario.tipo == usuario.Tipo.DESCRITOR:
        minhas = (
            Imagem.objects
            .filter(responsavel=usuario, ativo=True)
            .select_related("status")
        )
        disponiveis = minhas.filter(
            status__ativo=True,
            status__perfil_responsavel=Usuario.Tipo.DESCRITOR,
        ).filter(
            Q(status__permite_edicao=True)
            | Q(status__avanca_ao_abrir=True)
        )

        ctx.update({
            "visao": "descritor",
            "total_minhas": minhas.count(),
            "disponiveis": disponiveis,
            "disponiveis_count": disponiveis.count(),
        })

    elif usuario.tipo == usuario.Tipo.REVISOR:
        minhas = (
            Imagem.objects
            .filter(responsavel=usuario, ativo=True)
            .select_related("status")
        )
        para_conferir = minhas.filter(
            status__ativo=True,
            status__perfil_responsavel=Usuario.Tipo.REVISOR,
        ).filter(
            Q(status__permite_edicao=True)
            | Q(status__avanca_ao_abrir=True)
        )

        ctx.update({
            "visao": "revisor",
            "total_minhas": minhas.count(),
            "para_conferir": para_conferir,
            "para_conferir_count": para_conferir.count(),
        })

    return render(request, "core/dashboard.html", ctx)


# ============================================================
# MINHAS TAREFAS
# ============================================================

@login_required
def minhas_tarefas(request):
    from django.core.paginator import Paginator
    from django.db.models import Count
    from .models import filtro_autoria_imagem

    usuario = request.user
    status_slug = request.GET.get("status", "")
    busca = request.GET.get("busca", "").strip()
    pagina = request.GET.get("pagina", 1)
    lote_id = request.GET.get("lote", "")
    sem_lote = request.GET.get("sem_lote") == "1"

    # A lista vem do banco. Nenhum status é fixado pelo nome/slug no código.
    status_list = list(
        StatusWorkflow.objects
        .filter(ativo=True)
        .order_by("ordem")
    )
    slugs_visiveis = [s.slug for s in status_list]

    if usuario.tipo == usuario.Tipo.DESCRITOR:
        titulo_secao = "Minhas tarefas de descrição"
        acao_label = "Descrever"
        acao_icon = "bi-pencil-square"

    elif usuario.tipo == usuario.Tipo.REVISOR:
        titulo_secao = "Minhas tarefas de conferência"
        acao_label = "Conferir"
        acao_icon = "bi-eye"

    elif usuario.tipo in (usuario.Tipo.COORDENADOR, usuario.Tipo.ADMINISTRADOR):
        titulo_secao = "Todas as tarefas"
        acao_label = "Abrir"
        acao_icon = "bi-arrow-right-circle"

    else:
        titulo_secao = "Minhas tarefas"
        acao_label = "Abrir"
        acao_icon = "bi-arrow-right-circle"

    # Quando a coordenação abre o card "Imagens avulsas" pela tela de Lotes,
    # deixa explícito que esta listagem é o ponto de gestão das imagens sem lote.
    if (
        sem_lote
        and usuario.tipo in (
            usuario.Tipo.COORDENADOR,
            usuario.Tipo.ADMINISTRADOR,
        )
    ):
        titulo_secao = "Imagens avulsas"

    if usuario.tipo in (usuario.Tipo.COORDENADOR, usuario.Tipo.ADMINISTRADOR):
        tarefas = Imagem.objects.filter(
            ativo=True,
            status__ativo=True,
        )
    else:
        tarefas = (
            Imagem.objects
            .filter(
                ativo=True,
                status__ativo=True,
            )
            .filter(filtro_autoria_imagem(usuario))
            .distinct()
        )

    if status_slug:
        tarefas = tarefas.filter(status__slug=status_slug)

    if busca:
        tarefas = tarefas.filter(retranca__icontains=busca)

    lote_atual = None

    if sem_lote:
        tarefas = tarefas.filter(lote__isnull=True)

    elif lote_id:
        from .models import Lote

        lote_atual = Lote.objects.filter(pk=lote_id).first()
        if lote_atual:
            tarefas = tarefas.filter(lote=lote_atual)

    tarefas = (
        tarefas
        .select_related("status", "responsavel", "lote")
        .order_by("-criado_em")
    )

    base_contadores = Imagem.objects.filter(
        ativo=True,
        status__ativo=True,
    )

    if usuario.tipo not in (usuario.Tipo.COORDENADOR, usuario.Tipo.ADMINISTRADOR):
        base_contadores = (
            base_contadores
            .filter(filtro_autoria_imagem(usuario))
            .distinct()
        )

    if sem_lote:
        base_contadores = base_contadores.filter(lote__isnull=True)

    elif lote_atual:
        base_contadores = base_contadores.filter(lote=lote_atual)

    agregado = dict(
        base_contadores
        .values_list("status__slug")
        .annotate(qtd=Count("id"))
    )

    contadores = {
        status.slug: agregado.get(status.slug, 0)
        for status in status_list
    }

    # Usado apenas pela interface para escolher entre botão de ação e "Ver".
    # A lista é derivada das flags do banco, não de slugs fixos.
    perfil_operacional = _perfil_operacional(usuario)

    slugs_ativos = [
        status.slug
        for status in status_list
        if (
            (
                usuario.tipo == Usuario.Tipo.ADMINISTRADOR
                and (
                    status.permite_edicao
                    or status.avanca_ao_abrir
                    or status.revisao_concluida
                )
            )
            or (
                status.perfil_responsavel == perfil_operacional
                and (
                    status.permite_edicao
                    or status.avanca_ao_abrir
                    or status.revisao_concluida
                )
            )
        )
    ]

    total = tarefas.count()
    paginador = Paginator(tarefas, 50)
    pagina_obj = paginador.get_page(pagina)

    # ------------------------------------------------------------
    # Atribuição de imagens avulsas
    # ------------------------------------------------------------
    # A atribuição continua centralizada no fluxo de Lotes. Quando a
    # coordenação entra no card "Imagens avulsas", disponibilizamos os
    # responsáveis para atribuição individual de cada imagem.
    descritores = Usuario.objects.none()
    revisores = Usuario.objects.none()

    if (
        sem_lote
        and usuario.tipo in (
            usuario.Tipo.COORDENADOR,
            usuario.Tipo.ADMINISTRADOR,
        )
    ):
        descritores = (
            Usuario.objects
            .filter(
                tipo=Usuario.Tipo.DESCRITOR,
                is_active=True,
            )
            .order_by(
                "first_name",
                "last_name",
                "username",
            )
        )

        revisores = (
            Usuario.objects
            .filter(
                tipo=Usuario.Tipo.REVISOR,
                is_active=True,
            )
            .order_by(
                "first_name",
                "last_name",
                "username",
            )
        )

    # ------------------------------------------------------------
    # Navegação de retorno
    # ------------------------------------------------------------
    # Quando a tela foi aberta a partir de Lotes, o parâmetro ``next``
    # preserva exatamente aquela tela (inclusive filtros). Em acessos
    # antigos sem ``next``, lote/sem_lote continuam voltando para Lotes.
    url_voltar = _url_interna_segura(
        request,
        request.GET.get("next"),
    )

    # Para Descritor e Revisor, Minhas Tarefas faz parte do fluxo de Lotes.
    # Mesmo quando a tela é aberta diretamente pelo menu lateral (sem ``next``),
    # o botão Voltar deve levar para Lotes, e não para o Dashboard ou para a
    # própria tela de tarefas.
    if not url_voltar and usuario.tipo in (
        Usuario.Tipo.DESCRITOR,
        Usuario.Tipo.REVISOR,
    ):
        url_voltar = reverse("lotes_lista")

    # Coordenador/Admin ainda preservam a origem real quando disponível.
    if not url_voltar and (lote_atual or sem_lote):
        url_voltar = reverse("lotes_lista")

    if not url_voltar:
        url_voltar = _url_anterior_segura(request)

    if not url_voltar:
        url_voltar = reverse("dashboard")

    ctx = {
        "tarefas": pagina_obj,
        "pagina_obj": pagina_obj,
        "total": total,
        "titulo_secao": titulo_secao,
        "acao_label": acao_label,
        "acao_icon": acao_icon,
        "status_list": status_list,
        "status_slug_ativo": status_slug,
        "busca": busca,
        "contadores": contadores,
        "slugs_ativos": slugs_ativos,
        "mostrar_todos_status": usuario.tipo in (
            usuario.Tipo.COORDENADOR,
            usuario.Tipo.ADMINISTRADOR,
        ),
        "lote_atual": lote_atual,
        "sem_lote": sem_lote,
        "descritores": descritores,
        "revisores": revisores,
        "pode_atribuir_avulsas": (
            sem_lote
            and usuario.tipo in (
                usuario.Tipo.COORDENADOR,
                usuario.Tipo.ADMINISTRADOR,
            )
        ),
        "url_voltar": url_voltar,
    }

    return render(request, "core/minhas_tarefas.html", ctx)


# ============================================================
# IMAGENS — LISTAGEM
# ============================================================

@login_required
def imagens_lista(request):
    from django.core.paginator import Paginator
    from django.db.models import Count, Q
    from .models import Lote

    busca = request.GET.get("busca", "").strip()
    projeto_f = request.GET.get("projeto", "").strip()
    obra_f = request.GET.get("obra", "").strip()
    componente_f = request.GET.get("componente", "").strip()
    status_f = request.GET.get("status", "").strip()
    lote_f = request.GET.get("lote", "").strip()
    resp_descricao_f = request.GET.get("resp_descricao", "").strip()
    resp_revisao_f = request.GET.get("resp_revisao", "").strip()
    pagamento_descritor_f = request.GET.get("pagamento_descritor", "").strip()
    pagamento_revisor_f = request.GET.get("pagamento_revisor", "").strip()
    mostrar_inativas = request.GET.get("mostrar_inativas") == "1"
    pagina = request.GET.get("pagina", 1)

    # Quantidade de imagens por página.
    # O padrão é 25 para manter a listagem leve, mas o usuário pode
    # aumentar quando quiser visualizar mais registros de uma vez.
    por_pagina_opcoes = (25, 50, 100)

    try:
        por_pagina = int(
            request.GET.get(
                "por_pagina",
                25,
            )
        )
    except (TypeError, ValueError):
        por_pagina = 25

    if por_pagina not in por_pagina_opcoes:
        por_pagina = 25

    imagens = Imagem.objects.filter(ativo=False) if mostrar_inativas else Imagem.objects.filter(ativo=True)

    imagens = imagens.select_related(
        "status", "responsavel", "lote", "projeto", "componente_curricular",
        "descricao", "descricao__descritor", "descricao__revisor",
    )

    if busca:
        imagens = imagens.filter(retranca__icontains=busca)
    if projeto_f:
        imagens = imagens.filter(projeto__nome__icontains=projeto_f)
    if obra_f:
        imagens = imagens.filter(nome_obra__icontains=obra_f)
    if componente_f:
        imagens = imagens.filter(componente_curricular__nome__icontains=componente_f)
    if status_f:
        imagens = imagens.filter(status__nome__icontains=status_f)
    if resp_descricao_f:
        imagens = imagens.filter(
            Q(descricao__descritor__first_name__icontains=resp_descricao_f)
            | Q(descricao__descritor__last_name__icontains=resp_descricao_f)
            | Q(descricao__descritor__username__icontains=resp_descricao_f)
        )
    if resp_revisao_f:
        imagens = imagens.filter(
            Q(descricao__revisor__first_name__icontains=resp_revisao_f)
            | Q(descricao__revisor__last_name__icontains=resp_revisao_f)
            | Q(descricao__revisor__username__icontains=resp_revisao_f)
        )

    # Pagamento: converte o rótulo digitado ("Pago") para o valor interno ("pago")
    def _valor_pagamento(texto):
        texto_lower = texto.lower()
        for valor, label in Imagem.StatusPagamento.choices:
            if texto_lower == label.lower() or texto_lower == valor:
                return valor
        return None

    if pagamento_descritor_f:
        v = _valor_pagamento(pagamento_descritor_f)
        if v:
            imagens = imagens.filter(pagamento_descritor=v)
    if pagamento_revisor_f:
        v = _valor_pagamento(pagamento_revisor_f)
        if v:
            imagens = imagens.filter(pagamento_revisor=v)

    # Lote: match exato primeiro (nome completo selecionado no datalist), senão parcial
    lote_selecionado = None
    if lote_f:
        lote_selecionado = Lote.objects.filter(nome__iexact=lote_f, ativo=True).first()
        if not lote_selecionado:
            lote_selecionado = Lote.objects.filter(nome__icontains=lote_f, ativo=True).first()
        if lote_selecionado:
            imagens = imagens.filter(lote=lote_selecionado)

    imagens = imagens.order_by("-criado_em")

    total = imagens.count()
    paginador = Paginator(
        imagens,
        por_pagina,
    )
    pagina_obj = paginador.get_page(pagina)

    status_list = StatusWorkflow.objects.filter(ativo=True).order_by("ordem")
    descritores = Usuario.objects.filter(tipo=Usuario.Tipo.DESCRITOR, is_active=True).order_by("first_name", "username")
    revisores = Usuario.objects.filter(tipo=Usuario.Tipo.REVISOR, is_active=True).order_by("first_name", "username")

    lotes = (
        Lote.objects.filter(ativo=True)
        .annotate(total_ativo=Count("imagens", filter=Q(imagens__ativo=True)))
        .order_by("nome")
    )

    # Opções para os filtros de Projeto, Obra e Componente
    projetos_disponiveis = (
        Projeto.objects
        .filter(imagens__ativo=True)
        .distinct()
        .order_by("nome")
    )

    obras_disponiveis = (
        Imagem.objects.filter(ativo=True).exclude(nome_obra="")
        .values_list("nome_obra", flat=True).distinct().order_by()
    )
    componentes_disponiveis = (
        Imagem.objects.filter(ativo=True, componente_curricular__isnull=False)
        .values_list("componente_curricular__nome", flat=True).distinct().order_by()
    )

    # Preserva todos os filtros aplicados ao montar links de paginação
    querydict = request.GET.copy()
    querydict.pop("pagina", None)
    qs_sem_pagina = querydict.urlencode()

    tem_filtro = any([
        busca, projeto_f, obra_f, componente_f, status_f, lote_f,
        resp_descricao_f, resp_revisao_f,
        pagamento_descritor_f, pagamento_revisor_f, mostrar_inativas,
    ])

    ctx = {
        "imagens": pagina_obj,
        "status_list": status_list,
        "busca": busca,
        "projeto_f": projeto_f,
        "obra_f": obra_f,
        "componente_f": componente_f,
        "status_f": status_f,
        "lote_f": lote_f,
        "resp_descricao_f": resp_descricao_f,
        "resp_revisao_f": resp_revisao_f,
        "pagamento_descritor_f": pagamento_descritor_f,
        "pagamento_revisor_f": pagamento_revisor_f,
        "mostrar_inativas": mostrar_inativas,
        "tem_filtro": tem_filtro,
        "total": total,
        "pagina_obj": pagina_obj,
        "por_pagina": por_pagina,
        "por_pagina_opcoes": por_pagina_opcoes,
        "descritores": descritores,
        "revisores": revisores,
        "lotes": lotes,
        "lote_selecionado": lote_selecionado,
        "pode_atribuir_lote": _apenas_coordenador(request.user),
        "status_pagamento_choices": Imagem.StatusPagamento.choices,
        "projetos_disponiveis": projetos_disponiveis,
        "obras_disponiveis": sorted(set(obras_disponiveis)),
        "componentes_disponiveis": sorted(set(componentes_disponiveis)),
        "qs_sem_pagina": qs_sem_pagina,
    }
    return render(request, "core/imagens_lista.html", ctx)


# ============================================================
# IMAGENS — IMPORTAÇÃO VIA EXCEL
# ============================================================

@login_required
def importar_imagens(request):
    if not _apenas_coordenador(request.user):
        messages.error(
            request,
            "Apenas coordenadores e administradores podem importar imagens.",
        )
        return redirect("dashboard")

    from django.db.models.functions import Lower

    projetos_ativos = (
        Projeto.objects
        .filter(ativo=True)
        .order_by(Lower("nome"))
    )

    contexto_base = {
        "projetos": projetos_ativos,
        "projeto_selecionado_id": request.POST.get("projeto", ""),
        "acervo_fotoweb": request.POST.get("acervo_fotoweb", "").strip(),

        # Para o projeto atual, já deixamos os valores de teste preenchidos.
        # Em futuras importações eles podem ser alterados na própria tela.
        "projeto_editorial_pdf": (
            request.POST.get("projeto_editorial_pdf", "").strip()
            if request.method == "POST"
            else "G28_E002_EF2"
        ),
        "numero_projeto_pdf": (
            request.POST.get("numero_projeto_pdf", "").strip()
            if request.method == "POST"
            else "p013"
        ),
    }

    if request.method == "GET":
        return render(
            request,
            "core/importar_imagens.html",
            contexto_base,
        )

    # ------------------------------------------------------------
    # Projeto da importação
    # ------------------------------------------------------------
    projeto_id = request.POST.get("projeto", "").strip()

    if not projeto_id:
        messages.error(
            request,
            "Selecione o projeto ao qual este grupo de imagens pertence.",
        )
        return render(
            request,
            "core/importar_imagens.html",
            contexto_base,
        )

    try:
        projeto_selecionado = Projeto.objects.get(
            pk=projeto_id,
            ativo=True,
        )
    except (Projeto.DoesNotExist, ValueError):
        messages.error(
            request,
            "O projeto selecionado não existe ou está inativo.",
        )
        return render(
            request,
            "core/importar_imagens.html",
            contexto_base,
        )

    # ------------------------------------------------------------
    # Acervo FotoWeb da importação
    # ------------------------------------------------------------
    acervo_fotoweb = (
        request.POST.get("acervo_fotoweb", "")
        .strip()
        .strip("/")
    )

    if not acervo_fotoweb:
        messages.error(
            request,
            "Informe o Acervo FotoWeb deste relatório.",
        )
        return render(
            request,
            "core/importar_imagens.html",
            contexto_base,
        )

    # ------------------------------------------------------------
    # PDF no SharePoint
    # ------------------------------------------------------------
    projeto_editorial_pdf = (
        request.POST.get(
            "projeto_editorial_pdf",
            "",
        )
        .strip()
    )

    numero_projeto_pdf = (
        request.POST.get(
            "numero_projeto_pdf",
            "",
        )
        .strip()
    )

    projeto_editorial_pdf = (
        _normalizar_projeto_editorial_sharepoint(
            projeto_editorial_pdf
        )
    )

    numero_projeto_pdf = (
        _normalizar_numero_projeto_pdf(
            numero_projeto_pdf
        )
    )

    if not projeto_editorial_pdf:
        messages.error(
            request,
            "Informe o Projeto Editorial utilizado no SharePoint.",
        )
        return render(
            request,
            "core/importar_imagens.html",
            contexto_base,
        )

    if len(projeto_editorial_pdf.split("_")) < 2:
        messages.error(
            request,
            "Informe um Projeto Editorial válido. Ex.: G28_E002_EF2.",
        )
        return render(
            request,
            "core/importar_imagens.html",
            contexto_base,
        )

    if not numero_projeto_pdf:
        messages.error(
            request,
            "Informe o número do projeto Scriba do PDF. Ex.: p013.",
        )
        return render(
            request,
            "core/importar_imagens.html",
            contexto_base,
        )

    # ------------------------------------------------------------
    # Arquivo
    # ------------------------------------------------------------
    arquivo = request.FILES.get("arquivo")

    if not arquivo:
        messages.error(request, "Nenhum arquivo enviado.")
        return render(
            request,
            "core/importar_imagens.html",
            contexto_base,
        )

    if not arquivo.name.lower().endswith(".xlsx"):
        messages.error(
            request,
            "O arquivo deve estar no formato .xlsx",
        )
        return render(
            request,
            "core/importar_imagens.html",
            contexto_base,
        )

    status_inicial = _status_inicial_workflow()

    if not status_inicial:
        messages.error(
            request,
            "Nenhum status inicial ativo foi configurado "
            "no Gerenciador de Status.",
        )
        return render(
            request,
            "core/importar_imagens.html",
            contexto_base,
        )

    try:
        wb = openpyxl.load_workbook(
            arquivo,
            read_only=True,
            data_only=True,
        )
    except Exception as e:
        messages.error(
            request,
            f"Erro ao abrir o arquivo: {e}",
        )
        return render(
            request,
            "core/importar_imagens.html",
            contexto_base,
        )

    ws = wb.active
    headers = None
    data_rows = []

    for row in ws.iter_rows(values_only=True):
        if headers is None:
            headers = [
                _normalizar_cabecalho_excel(celula)
                for celula in row
            ]
            continue

        data_rows.append(
            dict(zip(headers, row))
        )

    wb.close()

    if not data_rows:
        messages.error(
            request,
            "O arquivo está vazio.",
        )
        return render(
            request,
            "core/importar_imagens.html",
            contexto_base,
        )

    # ------------------------------------------------------------
    # Validar colunas essenciais do relatório
    # ------------------------------------------------------------
    colunas_detectadas = set(headers or [])

    aliases_retranca = {
        "retranca",
    }

    aliases_obra = {
        "colecao",
        "obra",
        "nome_obra",
    }

    aliases_componente = {
        "disciplina",
        "componente",
        "componente_curricular",
    }

    if not colunas_detectadas.intersection(aliases_retranca):
        messages.error(
            request,
            "Não encontrei a coluna de Retranca no arquivo Excel.",
        )
        return render(
            request,
            "core/importar_imagens.html",
            contexto_base,
        )

    if not colunas_detectadas.intersection(aliases_obra):
        messages.error(
            request,
            "Não encontrei a coluna de Obra/Coleção no arquivo Excel.",
        )
        return render(
            request,
            "core/importar_imagens.html",
            contexto_base,
        )

    if not colunas_detectadas.intersection(aliases_componente):
        messages.error(
            request,
            "Não encontrei a coluna de Disciplina/Componente no arquivo Excel.",
        )
        return render(
            request,
            "core/importar_imagens.html",
            contexto_base,
        )

    importacao_id = uuid.uuid4()

    # ------------------------------------------------------------
    # Pré-carregar dados de referência
    # ------------------------------------------------------------

    # --------------------------------------------------------
    # Retrancas do arquivo e imagens já existentes
    # --------------------------------------------------------
    # Em vez de carregar TODAS as retrancas do banco para a memória,
    # consulta somente as retrancas presentes neste arquivo.
    retrancas_arquivo = {
        str(
            _valor_excel(
                linha,
                "retranca",
            )
            or ""
        ).strip()
        for linha in data_rows
    }
    retrancas_arquivo.discard("")

    imagens_existentes = {
        imagem.retranca: imagem
        for imagem in (
            Imagem.objects
            .filter(
                retranca__in=retrancas_arquivo
            )
            .only(
                "id",
                "retranca",
                "dados_fotoweb_originais",
            )
        )
    }

    retrancas_existentes = set(
        imagens_existentes
    )

    # Imagens duplicadas não são recriadas, mas recebem o snapshot
    # atualizado da linha original. Isso permite reimportar relatórios
    # antigos apenas para preparar a futura exportação FotoWeb.
    imagens_snapshot_atualizar = {}

    # Cache de usuários por username.
    usernames = {
        str(
            _valor_excel(
                r,
                "usuario",
                "responsavel",
            )
        ).strip()
        for r in data_rows
        if _valor_excel(
            r,
            "usuario",
            "responsavel",
        )
    }

    usuarios_cache = {
        u.username: u
        for u in Usuario.objects.filter(
            username__in=usernames,
        )
    }

    # Componentes já cadastrados são reaproveitados.
    # Se a planilha trouxer um componente novo, ele entra
    # automaticamente no Gerenciador de Componentes.
    componentes_cache = {
        componente.nome.casefold(): componente
        for componente in ComponenteCurricular.objects.all()
    }

    def resolver_componente(nome):
        nome = str(nome or "").strip()

        if not nome:
            return None

        chave = nome.casefold()
        componente = componentes_cache.get(chave)

        if componente:
            return componente

        componente, _ = (
            ComponenteCurricular.objects.get_or_create(
                nome=nome,
                defaults={"ativo": True},
            )
        )

        componentes_cache[chave] = componente
        return componente

    # Cache de idiomas.
    # O código ISO continua igual, mas o nome gravado passa a
    # ser localizado em português.
    idioma_cache = {}

    def resolver_idioma(lang_code):
        if lang_code in idioma_cache:
            return idioma_cache[lang_code]

        codigo_recebido = str(
            lang_code or ""
        ).strip()

        parte_idioma = (
            codigo_recebido
            .split("-")[0]
            .lower()
        )

        try:
            idioma = None

            if len(parte_idioma) == 2:
                idioma = pycountry.languages.get(
                    alpha_2=parte_idioma,
                )
            elif len(parte_idioma) == 3:
                idioma = pycountry.languages.get(
                    alpha_3=parte_idioma,
                )

            codigo = (
                idioma.alpha_3
                if idioma
                and hasattr(idioma, "alpha_3")
                else "und"
            )

            nome = _nome_idioma_pt(codigo)

        except Exception:
            codigo = "und"
            nome = "Idioma não identificado"

        idioma_cache[lang_code] = (
            codigo,
            nome,
        )

        return codigo, nome

    # ------------------------------------------------------------
    # Processar linhas
    # ------------------------------------------------------------
    imagens_criar = []
    rows_com_descricao = []
    puladas = 0
    erros = []
    log = []

    etapas_validas = [
        e[0]
        for e in Imagem.Etapa.choices
    ]

    for numero_linha, data in enumerate(
        data_rows,
        start=2,
    ):
        retranca = str(
            _valor_excel(
                data,
                "retranca",
            )
        ).strip()

        if not retranca:
            puladas += 1
            continue

        dados_fotoweb_originais = (
            _snapshot_linha_fotoweb(
                data,
                numero_linha=numero_linha,
                arquivo_origem=arquivo.name,
            )
        )

        if retranca in retrancas_existentes:
            puladas += 1

            imagem_existente = (
                imagens_existentes.get(
                    retranca
                )
            )

            if imagem_existente:
                imagem_existente.dados_fotoweb_originais = (
                    dados_fotoweb_originais
                )

                imagens_snapshot_atualizar[
                    imagem_existente.pk
                ] = imagem_existente

                mensagem_pulada = (
                    "Já existe — base original do FotoWeb atualizada; "
                    "projeto e workflow preservados"
                )
            else:
                # Duplicata dentro da própria planilha.
                mensagem_pulada = (
                    "Retranca repetida no arquivo — primeira ocorrência preservada"
                )

            log.append({
                "tipo": "pulada",
                "retranca": retranca,
                "msg": mensagem_pulada,
            })
            continue

        username = str(
            _valor_excel(
                data,
                "usuario",
                "responsavel",
            )
        ).strip()

        responsavel = usuarios_cache.get(
            username
        )

        etapa_raw = str(
            _valor_excel(
                data,
                "etapa",
            )
            or "AD"
        )

        etapa = (
            etapa_raw
            .replace("Etapa:", "")
            .strip()
        )

        if etapa not in etapas_validas:
            etapa = Imagem.Etapa.AD

        img_file = str(
            _valor_excel(
                data,
                "img_file",
                "arquivo",
                "caminho_arquivo",
            )
        )

        nome_arquivo = (
            os.path.basename(img_file)
            if img_file
            else ""
        )

        componente_raw = str(
            _valor_excel(
                data,
                "disciplina",
                "componente",
                "componente_curricular",
            )
        ).strip()

        volume_raw = str(
            _valor_excel(
                data,
                "volume",
                "volume_ano_modulo",
            )
        ).strip()

        url_pdf_sharepoint = (
            _montar_url_pdf_sharepoint(
                projeto_editorial_pdf,
                numero_projeto_pdf,
                componente_raw,
                volume_raw,
            )
        )

        imagem = Imagem(
            retranca=retranca,

            # Todas as imagens NOVAS desta importação
            # pertencem ao projeto escolhido no formulário.
            projeto=projeto_selecionado,

            nome_obra=str(
                _valor_excel(
                    data,
                    "colecao",
                    "obra",
                    "nome_obra",
                )
            ).strip(),
            componente_curricular=resolver_componente(
                componente_raw
            ),
            volume_ano_modulo=volume_raw,
            capitulo_unidade=str(
                _valor_excel(
                    data,
                    "capitulo",
                    "capitulo_unidade",
                )
            ).strip(),
            etapa=etapa,
            nome_arquivo=nome_arquivo,
            caminho_arquivo=img_file,

            # Link de pesquisa no FotoWeb montado automaticamente
            # a partir do Acervo desta importação + retranca.
            url_fotoweb=_montar_url_fotoweb(
                acervo_fotoweb,
                retranca,
            ),

            # O campo existente url_pdf passa a armazenar o link direto
            # do Manual do Professor no SharePoint.
            url_pdf=url_pdf_sharepoint,

            # Guarda os valores ORIGINAIS da linha do FotoWeb.
            # Na exportação específica, somente descricao e
            # descricao_flat serão substituídas.
            dados_fotoweb_originais=dados_fotoweb_originais,

            status=status_inicial,
            responsavel=responsavel,
            cadastrado_por=request.user,
            importacao_id=importacao_id,
            ativo=True,
        )

        imagens_criar.append(imagem)

        rows_com_descricao.append((
            retranca,
            responsavel,
            _valor_excel(
                data,
                "descricao",
            ),
        ))

        # Evita duplicatas dentro da própria planilha.
        retrancas_existentes.add(retranca)

        log.append({
            "tipo": "criada",
            "retranca": retranca,
            "msg": (
                f"Importada para o projeto "
                f"'{projeto_selecionado.nome}' "
                f"com FotoWeb preparado"
                + (
                    " e link do PDF no SharePoint gerado"
                    if url_pdf_sharepoint
                    else " — PDF sem dados suficientes para montar o link"
                )
            ),
        })

    # ------------------------------------------------------------
    # Bulk create imagens em lotes de 500
    # ------------------------------------------------------------
    LOTE = 500
    criadas = 0

    try:
        with transaction.atomic():

            # Atualiza somente o snapshot FotoWeb das imagens que já
            # existiam. Nenhum dado operacional do Dito é alterado.
            imagens_para_atualizar = list(
                imagens_snapshot_atualizar.values()
            )

            for i in range(
                0,
                len(imagens_para_atualizar),
                LOTE,
            ):
                Imagem.objects.bulk_update(
                    imagens_para_atualizar[
                        i:i + LOTE
                    ],
                    [
                        "dados_fotoweb_originais",
                    ],
                    batch_size=LOTE,
                )

            for i in range(
                0,
                len(imagens_criar),
                LOTE,
            ):
                lote = imagens_criar[
                    i:i + LOTE
                ]

                Imagem.objects.bulk_create(
                    lote,
                    ignore_conflicts=True,
                )

                criadas += len(lote)

            # Buscar IDs das imagens criadas.
            retrancas_criadas = [
                img.retranca
                for img in imagens_criar
            ]

            imagens_db = {
                img.retranca: img
                for img in Imagem.objects.filter(
                    retranca__in=retrancas_criadas
                )
            }

            # ----------------------------------------------------
            # Bulk create descrições
            # ----------------------------------------------------
            descricoes_criar = []
            trechos_por_retranca = {}

            for (
                retranca,
                responsavel,
                descricao_raw,
            ) in rows_com_descricao:

                imagem = imagens_db.get(
                    retranca
                )

                if (
                    not imagem
                    or not descricao_raw
                ):
                    continue

                try:
                    trechos_data = (
                        ast.literal_eval(
                            str(descricao_raw)
                        )
                    )

                    if (
                        not isinstance(
                            trechos_data,
                            list,
                        )
                        or not trechos_data
                    ):
                        continue

                except (
                    ValueError,
                    SyntaxError,
                ):
                    continue

                descricao = Descricao(
                    imagem=imagem,
                    descritor=responsavel,
                    descritor_bloqueado=False,
                )

                descricoes_criar.append(
                    descricao
                )

                trechos_por_retranca[
                    retranca
                ] = trechos_data

            for i in range(
                0,
                len(descricoes_criar),
                LOTE,
            ):
                Descricao.objects.bulk_create(
                    descricoes_criar[
                        i:i + LOTE
                    ],
                    ignore_conflicts=True,
                )

            # Buscar IDs das descrições criadas.
            descricoes_db = {
                d.imagem.retranca: d
                for d in Descricao.objects.filter(
                    imagem__retranca__in=list(
                        trechos_por_retranca.keys()
                    )
                ).select_related("imagem")
            }

            # ----------------------------------------------------
            # Bulk create trechos
            # ----------------------------------------------------
            trechos_criar = []

            for (
                retranca,
                trechos_data,
            ) in trechos_por_retranca.items():

                descricao = descricoes_db.get(
                    retranca
                )

                if not descricao:
                    continue

                for (
                    ordem,
                    trecho_data,
                ) in enumerate(
                    trechos_data,
                    1,
                ):
                    lang_code = (
                        trecho_data.get(
                            "lang",
                            "pt-BR",
                        )
                    )

                    texto = (
                        _normalizar_texto_descricao(
                            trecho_data.get(
                                "text",
                                "",
                            )
                        )
                    )

                    (
                        idioma_codigo,
                        idioma_nome,
                    ) = resolver_idioma(
                        lang_code
                    )

                    trechos_criar.append(
                        Trecho(
                            descricao=descricao,
                            ordem=ordem,
                            texto=texto,
                            idioma_codigo=(
                                idioma_codigo
                            ),
                            idioma_nome=(
                                idioma_nome
                            ),
                        )
                    )

            for i in range(
                0,
                len(trechos_criar),
                LOTE,
            ):
                Trecho.objects.bulk_create(
                    trechos_criar[
                        i:i + LOTE
                    ]
                )

    except Exception as e:
        erros.append({
            "retranca": "—",
            "msg": str(e),
        })
        criadas = 0

    ctx = {
        "resultado": True,
        "criadas": criadas,
        "puladas": puladas,
        "fotoweb_sincronizadas": len(
            imagens_snapshot_atualizar
        ),
        "erros": erros,

        # Limita o log para não travar o navegador.
        "log": log[:200],
        "log_truncado": len(log) > 200,
        "total_log": len(log),

        "importacao_id": (
            importacao_id
            if criadas > 0
            else None
        ),

        # Deixa explícito em qual Projeto, Acervo e configuração de rede
        # o grupo foi importado.
        "projeto_importacao": projeto_selecionado,
        "acervo_fotoweb": acervo_fotoweb,
        "projeto_editorial_pdf": projeto_editorial_pdf,
        "numero_projeto_pdf": numero_projeto_pdf,

        # Mantém a lista disponível caso seja necessário
        # renderizar novamente o formulário.
        "projetos": projetos_ativos,
        "projeto_selecionado_id": str(
            projeto_selecionado.pk
        ),
    }

    return render(
        request,
        "core/importar_imagens.html",
        ctx,
    )


# ============================================================
# DESCRIÇÃO DA IMAGEM
# ============================================================

@login_required
def descricao_imagem(request, pk):
    from .models import HistoricoItem

    imagem = get_object_or_404(
        Imagem.objects.select_related("status", "responsavel"),
        pk=pk,
        ativo=True,
    )
    usuario = request.user
    descricao = getattr(imagem, "descricao", None)

    # ------------------------------------------------------------
    # Navegação de retorno
    # ------------------------------------------------------------
    # A origem explícita (``next``) tem prioridade. Isso permite voltar
    # exatamente para o lote, para imagens avulsas ou para uma listagem
    # filtrada de tarefas. Se a URL foi aberta por um link antigo, usamos
    # o Referer interno como fallback.
    url_voltar = _url_interna_segura(
        request,
        request.GET.get("next"),
    )

    if not url_voltar:
        url_voltar = _url_anterior_segura(request)

    if not url_voltar:
        if usuario.tipo in (Usuario.Tipo.DESCRITOR, Usuario.Tipo.REVISOR):
            url_voltar = reverse("minhas_tarefas")
        else:
            url_voltar = reverse("imagens_lista")

    # ------------------------------------------------------------
    # Visualização
    # ------------------------------------------------------------
    pode_visualizar = _usuario_pode_visualizar_imagem(
        usuario,
        imagem,
        descricao,
    )

    if not pode_visualizar:
        return render(request, "core/descricao.html", {
            "imagem": imagem,
            "descricao": descricao,
            "trechos": [],
            "pode_editar": False,
            "pode_visualizar": False,
            "somente_leitura": True,
            "motivo_bloqueio": "Esta imagem não está disponível para você.",
            "idiomas_json": "[]",
            "lote_progresso_atual": None,
            "lote_total": None,
            "lote_eh_ultima": False,
            "url_voltar": url_voltar,
        })

    # ------------------------------------------------------------
    # Auto-transição ao abrir
    # ------------------------------------------------------------
    # O comportamento depende exclusivamente das flags do StatusWorkflow.
    # O nome e o slug do status não interferem na regra.
    if _usuario_pode_iniciar_status(usuario, imagem, descricao):
        status_anterior = imagem.status
        novo_status = status_anterior.proximo()

        if (
            novo_status
            and novo_status.ativo
            and novo_status.perfil_responsavel == status_anterior.perfil_responsavel
            and novo_status.permite_edicao
        ):
            perfil = status_anterior.perfil_responsavel
            tipo_acao = _tipo_acao_inicio(perfil)

            with transaction.atomic():
                if descricao:
                    campos_descricao = []

                    if (
                        perfil == Usuario.Tipo.DESCRITOR
                        and usuario.tipo == Usuario.Tipo.DESCRITOR
                        and descricao.descritor_id != usuario.id
                    ):
                        descricao.descritor = usuario
                        campos_descricao.append("descritor")

                    elif (
                        perfil == Usuario.Tipo.REVISOR
                        and usuario.tipo == Usuario.Tipo.REVISOR
                        and descricao.revisor_id != usuario.id
                    ):
                        descricao.revisor = usuario
                        campos_descricao.append("revisor")

                    elif (
                        perfil == Usuario.Tipo.COORDENADOR
                        and descricao.coordenador_id != usuario.id
                    ):
                        descricao.coordenador = usuario
                        campos_descricao.append("coordenador")

                    if campos_descricao:
                        descricao.save(update_fields=campos_descricao)

                imagem.status = novo_status

                # Na revisão final, o coordenador passa a ser o responsável
                # operacional pela imagem ao iniciar a etapa.
                if perfil == Usuario.Tipo.COORDENADOR:
                    imagem.responsavel = usuario
                    imagem.save(update_fields=["status", "responsavel"])
                else:
                    imagem.save(update_fields=["status"])

                HistoricoItem.objects.create(
                    imagem=imagem,
                    descricao=descricao,
                    usuario=usuario,
                    tipo_acao=tipo_acao,
                    status_anterior=status_anterior,
                    novo_status=novo_status,
                    observacao="Tarefa iniciada automaticamente ao abrir.",
                )

    # ------------------------------------------------------------
    # Permissão de edição
    # ------------------------------------------------------------
    pode_editar = _usuario_pode_editar_imagem(
        usuario,
        imagem,
        descricao,
    )

    # Uma revisão concluída não precisa voltar a ser editável para ser
    # finalizada. O coordenador/admin pode abrir a imagem em modo de
    # consulta e avançá-la individualmente para o status final.
    proximo_status = imagem.status.proximo() if imagem.status else None
    pode_finalizar = bool(
        pode_visualizar
        and imagem.status
        and imagem.status.revisao_concluida
        and proximo_status
        and proximo_status.is_final
        and (
            usuario.tipo == Usuario.Tipo.ADMINISTRADOR
            or imagem.status.perfil_responsavel == _perfil_operacional(usuario)
        )
    )

    somente_leitura = pode_visualizar and not pode_editar
    motivo_bloqueio = None

    if somente_leitura:
        if (
            usuario.tipo == Usuario.Tipo.DESCRITOR
            and descricao
            and descricao.descritor_id == usuario.id
            and descricao.descritor_bloqueado
        ):
            motivo_bloqueio = (
                "Você já enviou esta descrição. "
                "Somente o coordenador pode liberar novamente para edição."
            )

        elif (
            usuario.tipo == Usuario.Tipo.REVISOR
            and descricao
            and descricao.revisor_id == usuario.id
            and descricao.revisor_bloqueado
        ):
            motivo_bloqueio = (
                "Você já concluiu a conferência. "
                "Somente o coordenador pode liberar novamente para edição."
            )

        elif not pode_finalizar:
            motivo_bloqueio = (
                "Esta imagem está disponível apenas para consulta neste status."
            )

    # ------------------------------------------------------------
    # Trechos existentes
    # ------------------------------------------------------------
    trechos = []
    if descricao:
        trechos = descricao.trechos.filter(ativo=True).order_by("ordem")

    # ------------------------------------------------------------
    # Idiomas
    # ------------------------------------------------------------
    import json

    PRIORITARIOS = [
        "por", "eng", "spa", "fra", "deu", "ita", "jpn",
        "zho", "lat", "ara", "rus", "hin", "kor", "ell",
    ]

    idiomas_pt, nomes_idiomas = _catalogo_idiomas_pt()

    prioritarios = [
        item for item in idiomas_pt
        if item["codigo"] in PRIORITARIOS
    ]
    todos = [
        item for item in idiomas_pt
        if item["codigo"] not in PRIORITARIOS
    ]

    # Mantém a ordem definida acima para os idiomas mais usados e deixa
    # o restante em ordem alfabética pelo nome em português.
    ordem_prioritarios = {
        codigo: posicao
        for posicao, codigo in enumerate(PRIORITARIOS)
    }
    prioritarios.sort(
        key=lambda item: ordem_prioritarios.get(item["codigo"], 999)
    )
    todos.sort(key=lambda item: item["nome"].casefold())

    # Compatibilidade com descrições antigas. Se o banco ainda tiver um
    # nome em inglês, a tela usa sempre o nome localizado pelo código ISO.
    # Para códigos que não possuem nome pt-BR no CLDR, mostramos apenas
    # uma identificação neutra pelo código, nunca o nome inglês.
    for trecho in trechos:
        trecho.idioma_nome = nomes_idiomas.get(
            trecho.idioma_codigo,
            _nome_idioma_pt(trecho.idioma_codigo),
        )

    idiomas_json = json.dumps(
        prioritarios + todos,
        ensure_ascii=False,
    )

    # ------------------------------------------------------------
    # Progresso do usuário dentro do lote
    # ------------------------------------------------------------
    # Não existe uma "posição fixa" da imagem no lote. O usuário pode abrir
    # qualquer retranca primeiro. Por isso, o contador representa quantas
    # imagens ele já iniciou naquela etapa do workflow.
    #
    # Exemplo:
    # - abre qualquer imagem pela primeira vez -> 1/178
    # - abre outra imagem ainda não iniciada   -> 2/178
    # - reabre uma imagem já iniciada          -> continua 2/178
    lote_progresso_atual = None
    lote_total = None
    lote_eh_ultima = False

    if imagem.lote:
        escopo_lote = _escopo_fixo_do_lote(
            imagem.lote,
            usuario,
        )

        lote_total = escopo_lote.count()

        perfil_etapa = (
            imagem.status.perfil_responsavel
            if imagem.status
            else _perfil_operacional(usuario)
        )

        tipo_inicio = _tipo_acao_inicio(
            perfil_etapa
        )

        lote_progresso_atual = (
            HistoricoItem.objects
            .filter(
                imagem__in=escopo_lote,
                usuario=usuario,
                tipo_acao=tipo_inicio,
            )
            .values("imagem_id")
            .distinct()
            .count()
        )

        lote_eh_ultima = (
            _proxima_imagem_do_lote(
                imagem,
                usuario,
            ) is None
        )

    ctx = {
        "imagem": imagem,
        "descricao": descricao,
        "trechos": trechos,
        "pode_editar": pode_editar,
        "pode_visualizar": pode_visualizar,
        "pode_finalizar": pode_finalizar,
        "somente_leitura": somente_leitura,
        "motivo_bloqueio": motivo_bloqueio,
        "idiomas_json": idiomas_json,
        "lote_progresso_atual": lote_progresso_atual,
        "lote_total": lote_total,
        "lote_eh_ultima": lote_eh_ultima,
        "url_voltar": url_voltar,
    }

    return render(request, "core/descricao.html", ctx)


import json
from django.http import JsonResponse
from django.views.decorators.http import require_POST


@login_required
@require_POST
def salvar_trecho(request, pk):
    """Salva ou atualiza os trechos de uma descrição via AJAX."""
    from .models import HistoricoItem

    imagem = get_object_or_404(
        Imagem.objects.select_related("status", "responsavel"),
        pk=pk,
        ativo=True,
    )
    usuario = request.user
    descricao_atual = getattr(imagem, "descricao", None)

    if not _usuario_pode_editar_imagem(
        usuario,
        imagem,
        descricao_atual,
    ):
        return JsonResponse(
            {"ok": False, "erro": "Sem permissão para editar neste status."},
            status=403,
        )

    try:
        body = json.loads(request.body)
    except json.JSONDecodeError:
        return JsonResponse(
            {"ok": False, "erro": "JSON inválido."},
            status=400,
        )

    descricao, criada = Descricao.objects.get_or_create(
        imagem=imagem,
        defaults={
            "descritor": (
                usuario
                if usuario.tipo == Usuario.Tipo.DESCRITOR
                else None
            ),
            "descritor_bloqueado": False,
        },
    )

    # Mantém a autoria das etapas coerente mesmo quando a descrição já existia.
    campos_descricao = []

    if (
        usuario.tipo == Usuario.Tipo.DESCRITOR
        and descricao.descritor_id != usuario.id
    ):
        descricao.descritor = usuario
        campos_descricao.append("descritor")

    elif (
        usuario.tipo == Usuario.Tipo.REVISOR
        and descricao.revisor_id != usuario.id
    ):
        descricao.revisor = usuario
        campos_descricao.append("revisor")

    elif (
        usuario.tipo in (
            Usuario.Tipo.COORDENADOR,
            Usuario.Tipo.ADMINISTRADOR,
        )
        and imagem.status.perfil_responsavel == Usuario.Tipo.COORDENADOR
        and descricao.coordenador_id != usuario.id
    ):
        descricao.coordenador = usuario
        campos_descricao.append("coordenador")

    if campos_descricao:
        descricao.save(update_fields=campos_descricao)

    trechos_data = body.get("trechos", [])

    with transaction.atomic():
        descricao.trechos.all().delete()

        novos = []

        for i, t in enumerate(trechos_data, 1):
            texto = _normalizar_texto_descricao(
                t.get("texto", "")
            ).strip()
            idioma_codigo = t.get("idioma_codigo", "por")
            idioma_nome = t.get("idioma_nome", "Português")

            if texto:
                novos.append(
                    Trecho(
                        descricao=descricao,
                        ordem=i,
                        texto=texto,
                        idioma_codigo=idioma_codigo,
                        idioma_nome=idioma_nome,
                    )
                )

        Trecho.objects.bulk_create(novos)

        if usuario.tipo == Usuario.Tipo.DESCRITOR and criada:
            HistoricoItem.objects.create(
                imagem=imagem,
                descricao=descricao,
                usuario=usuario,
                tipo_acao=HistoricoItem.TipoAcao.DESCRICAO_INICIADA,
                observacao="Descrição iniciada pelo descritor.",
            )

    return JsonResponse({
        "ok": True,
        "total": len(novos),
    })


@login_required
@require_POST
def avancar_status(request, pk):
    """
    Avança a imagem para o próximo status ativo do workflow.

    As regras dependem das flags do StatusWorkflow, e não do nome/slug:
    - permite_edicao controla se a etapa pode ser concluída;
    - revisao_concluida permite o avanço final para o status is_final;
    - perfil_responsavel controla quem pode executar a transição;
    - flags de conclusão definem o tipo correto de histórico.
    """
    from .models import HistoricoItem

    imagem = get_object_or_404(
        Imagem.objects.select_related("status", "responsavel"),
        pk=pk,
        ativo=True,
    )
    usuario = request.user
    status_atual = imagem.status
    descricao = getattr(imagem, "descricao", None)

    if status_atual.is_final:
        return JsonResponse(
            {"ok": False, "erro": "Esta imagem já está finalizada."},
            status=400,
        )

    perfil_operacional = _perfil_operacional(usuario)

    if (
        usuario.tipo != Usuario.Tipo.ADMINISTRADOR
        and perfil_operacional != status_atual.perfil_responsavel
    ):
        return JsonResponse(
            {"ok": False, "erro": "Você não tem permissão para avançar esta tarefa."},
            status=403,
        )

    # Etapas normais só podem ser concluídas quando são editáveis.
    # A exceção é o marco "revisão concluída", que pode avançar para o
    # status final sem precisar ser editável.
    if not (
        status_atual.permite_edicao
        or status_atual.revisao_concluida
    ):
        return JsonResponse(
            {
                "ok": False,
                "erro": "Este status não permite avanço manual.",
            },
            status=400,
        )

    if (
        status_atual.exige_atribuicao
        and not imagem.responsavel
    ):
        return JsonResponse(
            {
                "ok": False,
                "erro": (
                    f"Atribua um(a) "
                    f"{status_atual.get_perfil_responsavel_display().lower()} "
                    "antes de avançar."
                ),
            },
            status=400,
        )

    if usuario.tipo in (Usuario.Tipo.DESCRITOR, Usuario.Tipo.REVISOR):
        if imagem.responsavel_id != usuario.id:
            return JsonResponse(
                {
                    "ok": False,
                    "erro": "Esta tarefa não está atribuída a você.",
                },
                status=403,
            )

    proximo_status = status_atual.proximo()

    if not proximo_status:
        return JsonResponse(
            {
                "ok": False,
                "erro": "Não há próximo status ativo configurado.",
            },
            status=400,
        )

    tipo_acao = _tipo_acao_transicao(
        status_atual,
        proximo_status,
    )
    mudou_perfil = (
        proximo_status.perfil_responsavel
        != status_atual.perfil_responsavel
    )

    with transaction.atomic():
        autoatribuido = False

        # --------------------------------------------------------
        # Bloqueio da etapa concluída
        # --------------------------------------------------------
        if mudou_perfil and descricao:
            if (
                status_atual.perfil_responsavel
                == Usuario.Tipo.DESCRITOR
            ):
                descricao.descritor_bloqueado = True
                descricao.save(
                    update_fields=["descritor_bloqueado"]
                )

                HistoricoItem.objects.create(
                    imagem=imagem,
                    descricao=descricao,
                    usuario=usuario,
                    tipo_acao=HistoricoItem.TipoAcao.DESCRITOR_BLOQUEADO,
                    status_anterior=status_atual,
                    novo_status=proximo_status,
                    observacao=(
                        "Acesso do descritor bloqueado "
                        "automaticamente após o envio."
                    ),
                )

            elif (
                status_atual.perfil_responsavel
                == Usuario.Tipo.REVISOR
            ):
                descricao.revisor_bloqueado = True
                descricao.save(
                    update_fields=["revisor_bloqueado"]
                )

                HistoricoItem.objects.create(
                    imagem=imagem,
                    descricao=descricao,
                    usuario=usuario,
                    tipo_acao=HistoricoItem.TipoAcao.REVISOR_BLOQUEADO,
                    status_anterior=status_atual,
                    novo_status=proximo_status,
                    observacao=(
                        "Acesso do revisor bloqueado "
                        "automaticamente após a conferência."
                    ),
                )

        # --------------------------------------------------------
        # Responsável da próxima fase
        # --------------------------------------------------------
        if mudou_perfil:
            if perfil_operacional == proximo_status.perfil_responsavel:
                imagem.responsavel = usuario
                autoatribuido = True
            else:
                imagem.responsavel = None

        imagem.status = proximo_status
        imagem.save()

        # --------------------------------------------------------
        # Autoria da revisão final
        # --------------------------------------------------------
        if (
            autoatribuido
            and proximo_status.perfil_responsavel == Usuario.Tipo.COORDENADOR
            and descricao
            and not descricao.coordenador_id
        ):
            descricao.coordenador = usuario
            descricao.save(update_fields=["coordenador"])

        # --------------------------------------------------------
        # Finalização
        # --------------------------------------------------------
        if proximo_status.is_final and descricao:
            descricao.finalizado = True
            descricao.save(update_fields=["finalizado"])

        HistoricoItem.objects.create(
            imagem=imagem,
            descricao=descricao,
            usuario=usuario,
            tipo_acao=tipo_acao,
            status_anterior=status_atual,
            novo_status=proximo_status,
        )

    # Mantém a origem durante a navegação sequencial do lote. Assim, ao
    # abrir a próxima imagem, o botão Voltar continua apontando para a
    # mesma tela de tarefas que originou o fluxo.
    url_retorno = _url_interna_segura(
        request,
        request.GET.get("next"),
    )

    proxima_url = None

    if imagem.lote:
        proxima = _proxima_imagem_do_lote(
            imagem,
            usuario,
        )

        if proxima:
            proxima_url = reverse(
                "descricao_imagem",
                kwargs={"pk": proxima.pk},
            )

            if url_retorno:
                proxima_url = (
                    f"{proxima_url}?"
                    f"{urlencode({'next': url_retorno})}"
                )

    return JsonResponse({
        "ok": True,
        "novo_status": proximo_status.nome,
        "novo_slug": proximo_status.slug,
        "proxima_url": proxima_url,
    })


# ============================================================
#  CRUD
# ============================================================

def _pode_gerenciar_imagem(usuario):
    """Regra de negócio: só Coordenador e Administrador gerenciam imagens."""
    return usuario.tipo in (Usuario.Tipo.ADMINISTRADOR, Usuario.Tipo.COORDENADOR)


@login_required
def imagem_criar(request):
    """Create — cadastra uma nova imagem individual."""
    if not _pode_gerenciar_imagem(request.user):
        messages.error(request, "Você não tem permissão para cadastrar imagens.")
        return redirect("imagens_lista")

    if request.method == "POST":
        form = ImagemForm(request.POST)
        if form.is_valid():
            imagem = form.save(commit=False)
            imagem.cadastrado_por = request.user
            imagem.save()
            messages.success(request, "Imagem cadastrada com sucesso.")
            return redirect("imagens_lista")
    else:
        # Pré-seleciona o status marcado como inicial no workflow.
        status_inicial = _status_inicial_workflow()
        form = ImagemForm(initial={"status": status_inicial})

    return render(request, "core/imagem_form.html", {
        "form": form,
        "titulo": "Cadastrar imagem",
    })


@login_required
def imagem_editar(request, pk):
    """Update — edita os dados de uma imagem existente."""
    imagem = get_object_or_404(Imagem, pk=pk, ativo=True)

    if not _pode_gerenciar_imagem(request.user):
        messages.error(request, "Você não tem permissão para editar imagens.")
        return redirect("imagens_lista")

    if request.method == "POST":
        form = ImagemForm(request.POST, instance=imagem)
        if form.is_valid():
            form.save()
            messages.success(request, "Imagem atualizada com sucesso.")
            return redirect("imagens_lista")
    else:
        form = ImagemForm(instance=imagem)

    return render(request, "core/imagem_form.html", {
        "form": form,
        "titulo": "Editar imagem",
        "imagem": imagem,
    })


@login_required
def imagem_excluir(request, pk):
    """Delete — exclusão LÓGICA (marca ativo=False, não apaga do banco)."""
    imagem = get_object_or_404(Imagem, pk=pk, ativo=True)

    if not _pode_gerenciar_imagem(request.user):
        messages.error(request, "Você não tem permissão para excluir imagens.")
        return redirect("imagens_lista")

    if request.method == "POST":
        imagem.ativo = False
        imagem.save()
        messages.success(request, "Imagem excluída com sucesso.")
        return redirect("imagens_lista")

    return render(request, "core/imagem_excluir.html", {
        "imagem": imagem,
    })


# ============================================================
# ATRIBUIÇÃO DE TAREFAS (Coordenador)
# ============================================================

@login_required
@require_POST
def atribuir_descritor(request, pk):
    """Coordenador atribui um Descritor ao status inicial de descrição."""
    from .models import HistoricoItem

    if not _apenas_coordenador(request.user):
        messages.error(
            request,
            "Você não tem permissão para atribuir tarefas.",
        )
        return redirect("imagens_lista")

    imagem = get_object_or_404(
        Imagem.objects.select_related("status"),
        pk=pk,
        ativo=True,
    )

    status_atual = imagem.status

    if not (
        status_atual.is_inicial
        and status_atual.ativo
        and status_atual.perfil_responsavel == Usuario.Tipo.DESCRITOR
        and status_atual.exige_atribuicao
    ):
        messages.error(
            request,
            "Esta imagem não está em um status inicial elegível para atribuição a descritor.",
        )
        return redirect(request.POST.get("next", "imagens_lista"))

    descritor_id = request.POST.get("descritor_id")
    descritor = Usuario.objects.filter(
        pk=descritor_id,
        tipo=Usuario.Tipo.DESCRITOR,
    ).first()

    if not descritor:
        messages.error(request, "Descritor inválido.")
        return redirect(request.POST.get("next", "imagens_lista"))

    if not descritor.contrato_ativo:
        messages.error(
            request,
            f"{descritor} não possui contrato ativo no momento.",
        )
        return redirect(request.POST.get("next", "imagens_lista"))

    with transaction.atomic():
        imagem.responsavel = descritor
        imagem.save(update_fields=["responsavel"])

        descricao = getattr(imagem, "descricao", None)
        if descricao:
            campos_descricao = []

            if descricao.descritor_id != descritor.id:
                descricao.descritor = descritor
                campos_descricao.append("descritor")

            if descricao.descritor_bloqueado:
                descricao.descritor_bloqueado = False
                campos_descricao.append("descritor_bloqueado")

            if campos_descricao:
                descricao.save(update_fields=campos_descricao)

        HistoricoItem.objects.create(
            imagem=imagem,
            descricao=descricao,
            usuario=request.user,
            tipo_acao=HistoricoItem.TipoAcao.TAREFA_ATRIBUIDA,
            status_anterior=imagem.status,
            novo_status=imagem.status,
            observacao=f"Atribuído ao descritor {descritor}.",
        )

    messages.success(
        request,
        f"Tarefa atribuída a {descritor}.",
    )
    return redirect(request.POST.get("next", "imagens_lista"))


@login_required
@require_POST
def liberar_conferencia(request, pk):
    """Coordenador libera uma descrição concluída para um Revisor."""
    from .models import HistoricoItem

    if not _apenas_coordenador(request.user):
        messages.error(
            request,
            "Você não tem permissão para liberar tarefas.",
        )
        return redirect("imagens_lista")

    imagem = get_object_or_404(
        Imagem.objects.select_related("status"),
        pk=pk,
        ativo=True,
    )

    if not imagem.status.descricao_concluida:
        messages.error(
            request,
            "Esta imagem ainda não está marcada como descrição concluída.",
        )
        return redirect(request.POST.get("next", "imagens_lista"))

    revisor_id = request.POST.get("revisor_id")
    revisor = Usuario.objects.filter(
        pk=revisor_id,
        tipo=Usuario.Tipo.REVISOR,
    ).first()

    if not revisor:
        messages.error(request, "Revisor inválido.")
        return redirect(request.POST.get("next", "imagens_lista"))

    if not revisor.contrato_ativo:
        messages.error(
            request,
            f"{revisor} não possui contrato ativo no momento.",
        )
        return redirect(request.POST.get("next", "imagens_lista"))

    proximo_status = _status_entrada_perfil(
        Usuario.Tipo.REVISOR,
        depois_de=imagem.status,
    )

    if not proximo_status:
        messages.error(
            request,
            "Não existe um status ativo de entrada para o Revisor após esta etapa.",
        )
        return redirect(request.POST.get("next", "imagens_lista"))

    status_anterior = imagem.status

    with transaction.atomic():
        imagem.responsavel = revisor
        imagem.status = proximo_status
        imagem.save(
            update_fields=[
                "responsavel",
                "status",
            ]
        )

        descricao = getattr(imagem, "descricao", None)
        if descricao:
            campos_descricao = []

            if descricao.revisor_id != revisor.id:
                descricao.revisor = revisor
                campos_descricao.append("revisor")

            if descricao.revisor_bloqueado:
                descricao.revisor_bloqueado = False
                campos_descricao.append("revisor_bloqueado")

            if campos_descricao:
                descricao.save(update_fields=campos_descricao)

        HistoricoItem.objects.create(
            imagem=imagem,
            descricao=descricao,
            usuario=request.user,
            tipo_acao=HistoricoItem.TipoAcao.LIBERADO_CONFERENCIA,
            status_anterior=status_anterior,
            novo_status=proximo_status,
            observacao=f"Liberado para o revisor {revisor}.",
        )

    messages.success(
        request,
        f"Tarefa liberada para {revisor}.",
    )
    return redirect(request.POST.get("next", "imagens_lista"))


@login_required
@require_POST
def devolver_descritor(request, pk):
    """Coordenador devolve a tarefa para o Descritor corrigir."""
    from .models import HistoricoItem

    if not _apenas_coordenador(request.user):
        messages.error(
            request,
            "Você não tem permissão para liberar tarefas.",
        )
        return redirect("imagens_lista")

    imagem = get_object_or_404(Imagem, pk=pk, ativo=True)
    descricao = getattr(imagem, "descricao", None)

    if not descricao or not descricao.descritor:
        messages.error(
            request,
            "Esta imagem ainda não possui um descritor definido.",
        )
        return redirect(request.POST.get("next", "imagens_lista"))

    observacao = request.POST.get("observacao", "").strip()

    status_liberado = _status_entrada_perfil(
        Usuario.Tipo.DESCRITOR,
    )

    if not status_liberado:
        messages.error(
            request,
            "Não existe um status ativo de entrada para o Descritor.",
        )
        return redirect(request.POST.get("next", "imagens_lista"))

    status_anterior = imagem.status

    with transaction.atomic():
        descricao.descritor_bloqueado = False
        descricao.save(
            update_fields=["descritor_bloqueado"]
        )

        imagem.status = status_liberado
        imagem.responsavel = descricao.descritor
        imagem.save(
            update_fields=[
                "status",
                "responsavel",
            ]
        )

        HistoricoItem.objects.create(
            imagem=imagem,
            descricao=descricao,
            usuario=request.user,
            tipo_acao=HistoricoItem.TipoAcao.DESCRITOR_LIBERADO,
            status_anterior=status_anterior,
            novo_status=status_liberado,
            observacao=(
                observacao
                or f"Devolvido ao descritor {descricao.descritor} para correção."
            ),
        )

    messages.success(
        request,
        f"Tarefa devolvida para o descritor {descricao.descritor}.",
    )
    return redirect(request.POST.get("next", "imagens_lista"))


@login_required
@require_POST
def devolver_revisor(request, pk):
    """Coordenador devolve a tarefa para o Revisor corrigir."""
    from .models import HistoricoItem

    if not _apenas_coordenador(request.user):
        messages.error(
            request,
            "Você não tem permissão para liberar tarefas.",
        )
        return redirect("imagens_lista")

    imagem = get_object_or_404(Imagem, pk=pk, ativo=True)
    descricao = getattr(imagem, "descricao", None)

    if not descricao or not descricao.revisor:
        messages.error(
            request,
            "Esta imagem ainda não possui um revisor definido.",
        )
        return redirect(request.POST.get("next", "imagens_lista"))

    observacao = request.POST.get("observacao", "").strip()

    status_liberado = _status_entrada_perfil(
        Usuario.Tipo.REVISOR,
    )

    if not status_liberado:
        messages.error(
            request,
            "Não existe um status ativo de entrada para o Revisor.",
        )
        return redirect(request.POST.get("next", "imagens_lista"))

    status_anterior = imagem.status

    with transaction.atomic():
        descricao.revisor_bloqueado = False
        descricao.save(
            update_fields=["revisor_bloqueado"]
        )

        imagem.status = status_liberado
        imagem.responsavel = descricao.revisor
        imagem.save(
            update_fields=[
                "status",
                "responsavel",
            ]
        )

        HistoricoItem.objects.create(
            imagem=imagem,
            descricao=descricao,
            usuario=request.user,
            tipo_acao=HistoricoItem.TipoAcao.REVISOR_LIBERADO,
            status_anterior=status_anterior,
            novo_status=status_liberado,
            observacao=(
                observacao
                or f"Devolvido ao revisor {descricao.revisor} para correção."
            ),
        )

    messages.success(
        request,
        f"Tarefa devolvida para o revisor {descricao.revisor}.",
    )
    return redirect(request.POST.get("next", "imagens_lista"))


# ============================================================
# ATRIBUIÇÃO POR LOTE INTEIRO (Coordenador)
# ============================================================

@login_required
@require_POST
def atribuir_lote(request, lote_id):
    """
    Atribui imagens elegíveis de um lote a um Descritor ou Revisor.

    A elegibilidade é definida pelas flags do StatusWorkflow, nunca pelo
    nome/slug do status.
    """
    from .models import Lote, HistoricoItem

    if not _apenas_coordenador(request.user):
        messages.error(
            request,
            "Você não tem permissão para atribuir lotes.",
        )
        return redirect(request.POST.get("next", "imagens_lista"))

    lote = get_object_or_404(Lote, pk=lote_id, ativo=True)

    tipo_acao = request.POST.get("tipo_acao")
    usuario_id = request.POST.get("usuario_id")

    if tipo_acao == "descritor":
        usuario_alvo = Usuario.objects.filter(
            pk=usuario_id,
            tipo=Usuario.Tipo.DESCRITOR,
        ).first()

        imagens_elegiveis_qs = Imagem.objects.filter(
            lote=lote,
            ativo=True,
            status__ativo=True,
            status__is_inicial=True,
            status__perfil_responsavel=Usuario.Tipo.DESCRITOR,
            status__exige_atribuicao=True,
        )

    elif tipo_acao == "revisor":
        usuario_alvo = Usuario.objects.filter(
            pk=usuario_id,
            tipo=Usuario.Tipo.REVISOR,
        ).first()

        imagens_elegiveis_qs = Imagem.objects.filter(
            lote=lote,
            ativo=True,
            status__ativo=True,
            status__descricao_concluida=True,
        )

    else:
        messages.error(request, "Ação inválida.")
        return redirect(request.POST.get("next", "imagens_lista"))

    if not usuario_alvo:
        messages.error(
            request,
            "Usuário inválido para essa ação.",
        )
        return redirect(request.POST.get("next", "imagens_lista"))

    if not usuario_alvo.contrato_ativo:
        messages.error(
            request,
            f"{usuario_alvo} não possui contrato ativo no momento.",
        )
        return redirect(request.POST.get("next", "imagens_lista"))

    imagens_lote = Imagem.objects.filter(
        lote=lote,
        ativo=True,
    )
    imagens_elegiveis = list(
        imagens_elegiveis_qs.select_related("status")
    )

    total_lote = imagens_lote.count()
    total_elegiveis = len(imagens_elegiveis)
    ignoradas = total_lote - total_elegiveis

    if total_elegiveis == 0:
        messages.error(
            request,
            f"Nenhuma imagem do lote '{lote.nome}' está elegível para essa ação.",
        )
        return redirect(request.POST.get("next", "imagens_lista"))

    with transaction.atomic():
        if tipo_acao == "descritor":
            for imagem in imagens_elegiveis:
                imagem.responsavel = usuario_alvo
                imagem.save(
                    update_fields=["responsavel"]
                )

                descricao = getattr(imagem, "descricao", None)
                if descricao:
                    campos_descricao = []

                    if descricao.descritor_id != usuario_alvo.id:
                        descricao.descritor = usuario_alvo
                        campos_descricao.append("descritor")

                    if descricao.descritor_bloqueado:
                        descricao.descritor_bloqueado = False
                        campos_descricao.append("descritor_bloqueado")

                    if campos_descricao:
                        descricao.save(update_fields=campos_descricao)

                HistoricoItem.objects.create(
                    imagem=imagem,
                    descricao=descricao,
                    usuario=request.user,
                    tipo_acao=HistoricoItem.TipoAcao.TAREFA_ATRIBUIDA,
                    status_anterior=imagem.status,
                    novo_status=imagem.status,
                    observacao=(
                        f"Atribuído ao descritor {usuario_alvo} "
                        f"via lote '{lote.nome}'."
                    ),
                )

        elif tipo_acao == "revisor":
            for imagem in imagens_elegiveis:
                proximo_status = _status_entrada_perfil(
                    Usuario.Tipo.REVISOR,
                    depois_de=imagem.status,
                )

                if not proximo_status:
                    continue

                status_anterior = imagem.status
                imagem.responsavel = usuario_alvo
                imagem.status = proximo_status
                imagem.save(
                    update_fields=[
                        "responsavel",
                        "status",
                    ]
                )

                descricao = getattr(imagem, "descricao", None)
                if descricao:
                    campos_descricao = []

                    if descricao.revisor_id != usuario_alvo.id:
                        descricao.revisor = usuario_alvo
                        campos_descricao.append("revisor")

                    if descricao.revisor_bloqueado:
                        descricao.revisor_bloqueado = False
                        campos_descricao.append("revisor_bloqueado")

                    if campos_descricao:
                        descricao.save(update_fields=campos_descricao)

                HistoricoItem.objects.create(
                    imagem=imagem,
                    descricao=descricao,
                    usuario=request.user,
                    tipo_acao=HistoricoItem.TipoAcao.LIBERADO_CONFERENCIA,
                    status_anterior=status_anterior,
                    novo_status=proximo_status,
                    observacao=(
                        f"Liberado para o revisor {usuario_alvo} "
                        f"via lote '{lote.nome}'."
                    ),
                )

    msg = (
        f"{total_elegiveis} imagem(ns) do lote '{lote.nome}' "
        f"atribuída(s) a {usuario_alvo}."
    )

    if ignoradas:
        msg += (
            f" {ignoradas} imagem(ns) ignorada(s) "
            "por não estarem elegíveis para essa ação."
        )

    messages.success(request, msg)
    return redirect(request.POST.get("next", "imagens_lista"))


# ============================================================
# ORGANIZAÇÃO EM LOTES (pós-importação)
# ============================================================

@login_required
def organizar_lotes(request, importacao_id=None):
    """
    Tela de organização em Lotes. Funciona de duas formas:
    - Com importacao_id: mostra só as imagens daquela importação específica
      sem lote (fluxo logo após o upload do .xlsx).
    - Sem importacao_id: mostra todas as imagens ativas sem lote no sistema
      (acesso geral via botão "Criar Lote" na listagem de Imagens).
    """
    from .models import Lote

    if not _apenas_coordenador(request.user):
        messages.error(request, "Você não tem permissão para organizar lotes.")
        return redirect("dashboard")

    if importacao_id:
        imagens_sem_lote = Imagem.objects.filter(
            importacao_id=importacao_id, lote__isnull=True, ativo=True,
        )
    else:
        imagens_sem_lote = Imagem.objects.filter(lote__isnull=True, ativo=True)

    def _redirect_organizar_lotes():
        if importacao_id:
            return redirect("organizar_lotes", importacao_id=importacao_id)
        return redirect("organizar_lotes_geral")

    if request.method == "POST":
        nome_lote = request.POST.get("nome_lote", "").strip()
        descricao_lote = request.POST.get("descricao_lote", "").strip()
        data_prevista_raw = request.POST.get("data_prevista", "").strip()
        imagem_ids = request.POST.getlist("imagem_ids")

        try:
            data_prevista = _data_iso_ou_none(data_prevista_raw)
        except ValueError:
            messages.error(request, "Informe uma data prevista válida.")
            return _redirect_organizar_lotes()

        if not nome_lote:
            messages.error(request, "Informe um nome para o lote.")
        elif not imagem_ids:
            messages.error(request, "Selecione ao menos uma imagem para formar o lote.")
        elif Lote.objects.filter(nome=nome_lote).exists():
            messages.error(request, f"Já existe um lote com o nome '{nome_lote}'. Escolha outro nome.")
        else:
            lote = Lote.objects.create(
                nome=nome_lote,
                descricao=descricao_lote,
                data_prevista=data_prevista,
                criado_por=request.user,
            )
            atualizadas = imagens_sem_lote.filter(pk__in=imagem_ids).update(lote=lote)
            lote.sincronizar_data_efetiva()
            messages.success(request, f"Lote '{nome_lote}' criado com {atualizadas} imagem(ns).")

        return _redirect_organizar_lotes()

    # ---- Filtros (GET) ----
    componente = request.GET.get("componente", "").strip()
    volume = request.GET.get("volume", "").strip()
    busca = request.GET.get("busca", "").strip()

    imagens = imagens_sem_lote.select_related("status", "componente_curricular")

    if componente:
        imagens = imagens.filter(componente_curricular__nome__icontains=componente)
    if volume:
        imagens = imagens.filter(volume_ano_modulo__icontains=volume)
    if busca:
        imagens = imagens.filter(retranca__icontains=busca)

    imagens = imagens.order_by("componente_curricular__nome", "volume_ano_modulo", "retranca")

    componentes_disponiveis = (
        imagens_sem_lote.filter(componente_curricular__isnull=False)
        .values_list("componente_curricular__nome", flat=True)
        .distinct().order_by()
    )
    volumes_disponiveis = (
        imagens_sem_lote.exclude(volume_ano_modulo="")
        .values_list("volume_ano_modulo", flat=True)
        .distinct().order_by()
    )

    if importacao_id:
        lotes_recentes = Lote.objects.filter(
            imagens__importacao_id=importacao_id
        ).distinct().order_by("-criado_em")
    else:
        lotes_recentes = Lote.objects.filter(ativo=True).order_by("-criado_em")[:10]

    ctx = {
        "importacao_id": importacao_id,
        "imagens": imagens,
        "total_restante": imagens_sem_lote.count(),
        "componente": componente,
        "volume": volume,
        "busca": busca,
        "componentes_disponiveis": sorted(set(componentes_disponiveis)),
        "volumes_disponiveis": sorted(set(volumes_disponiveis)),
        "lotes_desta_importacao": lotes_recentes,
    }
    return render(request, "core/organizar_lotes.html", ctx)

# ============================================================
# ATUALIZAÇÃO DE STATUS DE PAGAMENTO
# ============================================================

@login_required
@require_POST
def atualizar_pagamento(request, pk):
    """
    Atualiza o status de pagamento (Contabilizado/Pago) do descritor
    ou do revisor de uma imagem. Restrito a Coordenador/Administrador.
    """
    from .models import HistoricoItem

    if not _apenas_coordenador(request.user):
        messages.error(request, "Você não tem permissão para atualizar pagamentos.")
        return redirect(request.POST.get("next", "imagens_lista"))

    imagem = get_object_or_404(Imagem, pk=pk, ativo=True)

    tipo = request.POST.get("tipo")  # "descritor" ou "revisor"
    novo_status = request.POST.get("status_pagamento")  # "contabilizado" ou "pago"

    valores_validos = [c[0] for c in Imagem.StatusPagamento.choices]
    if tipo not in ("descritor", "revisor") or novo_status not in valores_validos:
        messages.error(request, "Dados inválidos para atualização de pagamento.")
        return redirect(request.POST.get("next", "imagens_lista"))

    campo = f"pagamento_{tipo}"
    status_anterior_valor = getattr(imagem, campo)

    if status_anterior_valor == novo_status:
        messages.info(request, "O status de pagamento já estava definido dessa forma.")
        return redirect(request.POST.get("next", "imagens_lista"))

    setattr(imagem, campo, novo_status)
    imagem.save()

    HistoricoItem.objects.create(
        imagem=imagem,
        descricao=getattr(imagem, "descricao", None),
        usuario=request.user,
        tipo_acao=HistoricoItem.TipoAcao.PAGAMENTO_ATUALIZADO,
        observacao=(
            f"Pagamento do {tipo} alterado de "
            f"'{status_anterior_valor}' para '{novo_status}'."
        ),
    )

    messages.success(request, f"Pagamento do {tipo} atualizado para '{novo_status}'.")
    return redirect(request.POST.get("next", "imagens_lista"))


# ============================================================
# EDIÇÃO DE LOTE
# ============================================================

@login_required
def lote_editar(request, pk):
    """
    Edita os dados gerais do lote e permite administrar sua composição.

    A movimentação de imagens entre lotes altera apenas a FK ``Imagem.lote``:
    status, responsável, descrição, pagamentos e histórico editorial permanecem
    intactos. Remover uma imagem do lote a transforma em imagem avulsa.
    """
    from .models import Lote

    if not _apenas_coordenador(request.user):
        messages.error(request, "Você não tem permissão para editar lotes.")
        return redirect("imagens_lista")

    lote = get_object_or_404(Lote, pk=pk)

    def contexto_formulario():
        return {
            "lote": lote,
            "imagens_lote": (
                lote.imagens
                .filter(ativo=True)
                .select_related("status", "responsavel")
                .order_by("retranca")
            ),
            "lotes_destino": (
                Lote.objects
                .filter(ativo=True)
                .exclude(pk=lote.pk)
                .order_by("nome")
            ),
        }

    if request.method == "POST":
        # ------------------------------------------------------------
        # Gestão das imagens que pertencem ao lote
        # ------------------------------------------------------------
        acao_imagens = request.POST.get("acao_imagens", "").strip()

        if acao_imagens in ("mover", "remover"):
            ids_selecionados = request.POST.getlist("imagens_selecionadas")

            imagens = lote.imagens.filter(
                ativo=True,
                pk__in=ids_selecionados,
            )
            quantidade = imagens.count()

            if quantidade == 0:
                messages.warning(request, "Selecione pelo menos uma imagem do lote.")
                return redirect("lote_editar", pk=lote.pk)

            if acao_imagens == "remover":
                with transaction.atomic():
                    imagens.update(lote=None)
                    lote.sincronizar_data_efetiva()

                messages.success(
                    request,
                    f"{quantidade} imagem(ns) removida(s) do lote e deixada(s) como avulsa(s).",
                )
                return redirect("lote_editar", pk=lote.pk)

            destino_id = request.POST.get("lote_destino", "").strip()
            if not destino_id:
                messages.warning(request, "Escolha o lote de destino.")
                return redirect("lote_editar", pk=lote.pk)

            destino = get_object_or_404(
                Lote,
                pk=destino_id,
                ativo=True,
            )

            if destino.pk == lote.pk:
                messages.warning(request, "Escolha um lote de destino diferente do lote atual.")
                return redirect("lote_editar", pk=lote.pk)

            with transaction.atomic():
                imagens.update(lote=destino)
                lote.sincronizar_data_efetiva()
                destino.sincronizar_data_efetiva()

            messages.success(
                request,
                f"{quantidade} imagem(ns) movida(s) de '{lote.nome}' para '{destino.nome}'.",
            )
            return redirect("lote_editar", pk=lote.pk)

        # ------------------------------------------------------------
        # Dados gerais do lote
        # ------------------------------------------------------------
        nome = request.POST.get("nome", "").strip()
        descricao = request.POST.get("descricao", "").strip()
        data_prevista_raw = request.POST.get("data_prevista", "").strip()
        data_efetiva_raw = request.POST.get("data_efetiva", "").strip()
        ativo = request.POST.get("ativo") == "on"

        try:
            data_prevista = _data_iso_ou_none(data_prevista_raw)
            data_efetiva = _data_iso_ou_none(data_efetiva_raw)
        except ValueError:
            messages.error(request, "Confira as datas informadas e tente novamente.")
            return render(request, "core/lote_form.html", contexto_formulario())

        if not nome:
            messages.error(request, "O nome do lote não pode ficar vazio.")
        elif Lote.objects.exclude(pk=lote.pk).filter(nome=nome).exists():
            messages.error(request, f"Já existe outro lote com o nome '{nome}'.")
        else:
            lote.nome = nome
            lote.descricao = descricao
            lote.data_prevista = data_prevista
            lote.data_efetiva = data_efetiva
            lote.ativo = ativo
            lote.save()
            lote.sincronizar_data_efetiva()

            messages.success(request, f"Lote '{lote.nome}' atualizado com sucesso.")
            return redirect("lotes_lista")

    return render(request, "core/lote_form.html", contexto_formulario())

# ============================================================
# LISTAGEM DE LOTES
# ============================================================

@login_required
def lotes_lista(request):
    """
    Tela principal de Lotes. Comportamento por perfil:
    - Coordenador/Administrador: todos os lotes, com progresso completo por status
      e formulário de atribuição.
    - Descritor/Revisor: apenas os lotes onde ele tem imagens atribuídas,
      com progresso pessoal simplificado.
    Ambos veem um card de "Imagens avulsas" (sem lote), quando houver.
    """
    from .models import Lote, filtro_autoria_imagem

    usuario = request.user
    eh_coordenacao = _apenas_coordenador(usuario)

    busca = request.GET.get("busca", "").strip()
    mostrar_inativos = request.GET.get("mostrar_inativos") == "1"

    if eh_coordenacao:
        lotes = Lote.objects.filter(ativo=not mostrar_inativos)
    else:
        # Lotes onde o usuário é responsável atual OU foi autor da fase dele
        # (mesmo depois de já ter entregue — para o lote continuar visível
        # em modo consulta após o handoff).
        lotes = Lote.objects.filter(
            ativo=True,
            imagens__ativo=True,
        ).filter(
            imagens__in=Imagem.objects.filter(filtro_autoria_imagem(usuario))
        ).distinct()

    if busca:
        lotes = lotes.filter(nome__icontains=busca)

    lotes = lotes.select_related("criado_por").order_by("-criado_em")

    lotes_dados = []
    for lote in lotes:
        if eh_coordenacao:
            lotes_dados.append({
                "obj": lote,
                "total": lote.total_imagens,
                "progresso": lote.progresso_por_status(),
                "meu_progresso": None,
            })
        else:
            meu = lote.progresso_do_usuario(usuario)
            if meu:
                lotes_dados.append({
                    "obj": lote,
                    "total": meu["total"],
                    "progresso": None,
                    "meu_progresso": meu,
                })

    # ---- Card de imagens avulsas (sem lote) ----
    avulsas_qs = Imagem.objects.filter(ativo=True, lote__isnull=True)
    if not eh_coordenacao:
        avulsas_qs = avulsas_qs.filter(filtro_autoria_imagem(usuario)).distinct()
    total_avulsas = avulsas_qs.count()

    descritores = Usuario.objects.filter(
        tipo=Usuario.Tipo.DESCRITOR, is_active=True
    ).order_by("first_name", "username")

    revisores = Usuario.objects.filter(
        tipo=Usuario.Tipo.REVISOR, is_active=True
    ).order_by("first_name", "username")

    status_disponiveis = StatusWorkflow.objects.filter(ativo=True).order_by("ordem")

    ctx = {
        "lotes_dados": lotes_dados,
        "busca": busca,
        "mostrar_inativos": mostrar_inativos,
        "total_lotes": len(lotes_dados),
        "total_avulsas": total_avulsas,
        "eh_coordenacao": eh_coordenacao,
        "descritores": descritores,
        "revisores": revisores,
        "status_disponiveis": status_disponiveis,
    }
    return render(request, "core/lotes_lista.html", ctx)
# ============================================================
# NAVEGAÇÃO SEQUENCIAL DENTRO DO LOTE
# ============================================================

@login_required
def proxima_imagem_lote(request, pk):
    """
    Leva o usuário para a próxima imagem pendente do mesmo lote.
    Se não houver mais, informa e volta para a listagem apropriada.
    """
    imagem_atual = get_object_or_404(Imagem, pk=pk, ativo=True)

    if not imagem_atual.lote:
        messages.info(request, "Esta imagem não pertence a um lote.")
        return redirect("minhas_tarefas")

    proxima = _proxima_imagem_do_lote(imagem_atual, request.user)

    if proxima:
        return redirect("descricao_imagem", pk=proxima.pk)

    messages.success(
        request,
        f"Não há mais imagens pendentes para você no lote '{imagem_atual.lote.nome}'."
    )

    if _apenas_coordenador(request.user):
        return redirect(f"{reverse('imagens_lista')}?lote={imagem_atual.lote.nome}")
    return redirect("minhas_tarefas")

# ============================================================
# DEVOLUÇÃO DO LOTE AO COORDENADOR
# ============================================================

@login_required
@require_POST
def devolver_lote(request, pk):
    """
    Endpoint legado mantido temporariamente para compatibilidade com abas
    antigas abertas no navegador. A conclusão por lote foi desativada:
    cada imagem deve avançar individualmente pelo endpoint avancar_status.
    """
    get_object_or_404(Imagem, pk=pk, ativo=True)

    return JsonResponse(
        {
            "ok": False,
            "erro": (
                "A conclusão do lote inteiro foi desativada. "
                "Atualize a página e conclua cada imagem individualmente."
            ),
        },
        status=409,
    )


# ============================================================
# SOLICITAÇÃO PÚBLICA DE ACESSO
# ============================================================

def solicitar_acesso(request):
    """
    Tela pública: qualquer pessoa pode solicitar acesso ao Dito!.
    O usuário é criado como PENDENTE e inativo. O administrador
    aprova e define o perfil posteriormente.
    """
    from .forms import SolicitacaoAcessoForm

    if request.user.is_authenticated:
        return redirect("dashboard")

    if request.method == "POST":
        form = SolicitacaoAcessoForm(request.POST)
        if form.is_valid():
            form.save()
            return render(request, "registration/solicitacao_enviada.html")
    else:
        form = SolicitacaoAcessoForm()

    return render(request, "registration/solicitar_acesso.html", {"form": form})

# ============================================================
# GESTÃO DE USUÁRIOS (Administrador)
# ============================================================

def _apenas_admin(user):
    return user.is_authenticated and user.tipo == Usuario.Tipo.ADMINISTRADOR


@login_required
def usuarios_lista(request):
    """
    Tela de gestão de usuários, restrita ao Administrador.
    Mostra as solicitações pendentes no topo e os usuários ativos abaixo.
    """
    if not _apenas_admin(request.user):
        messages.error(request, "Apenas administradores podem acessar a gestão de usuários.")
        return redirect("dashboard")

    pendentes = Usuario.objects.filter(
        situacao=Usuario.Situacao.PENDENTE
    ).order_by("date_joined")

    usuarios = Usuario.objects.filter(
        situacao=Usuario.Situacao.APROVADO
    ).order_by("first_name", "email")

    ctx = {
        "pendentes": pendentes,
        "usuarios": usuarios,
        "total_pendentes": pendentes.count(),
        "tipos": Usuario.Tipo.choices,
    }
    return render(request, "core/usuarios_lista.html", ctx)


@login_required
@require_POST
def aprovar_solicitacao(request, pk):
    """Aprova uma solicitação, definindo o perfil e ativando o acesso."""
    from django.utils import timezone

    if not _apenas_admin(request.user):
        messages.error(request, "Apenas administradores podem aprovar solicitações.")
        return redirect("dashboard")

    solicitante = get_object_or_404(Usuario, pk=pk, situacao=Usuario.Situacao.PENDENTE)
    tipo = request.POST.get("tipo")

    tipos_validos = [t[0] for t in Usuario.Tipo.choices]
    if tipo not in tipos_validos:
        messages.error(request, "Selecione um perfil válido para aprovar.")
        return redirect("usuarios_lista")

    solicitante.tipo = tipo
    solicitante.situacao = Usuario.Situacao.APROVADO
    solicitante.is_active = True
    solicitante.aprovado_por = request.user
    solicitante.decidido_em = timezone.now()
    solicitante.save()

    messages.success(
        request,
        f"{solicitante.get_full_name() or solicitante.email} aprovado(a) como "
        f"{solicitante.get_tipo_display()}."
    )
    return redirect("usuarios_lista")


@login_required
@require_POST
def recusar_solicitacao(request, pk):
    """Recusa uma solicitação. O registro é mantido, mas o acesso não é liberado."""
    from django.utils import timezone

    if not _apenas_admin(request.user):
        messages.error(request, "Apenas administradores podem recusar solicitações.")
        return redirect("dashboard")

    solicitante = get_object_or_404(Usuario, pk=pk, situacao=Usuario.Situacao.PENDENTE)
    observacao = request.POST.get("observacao", "").strip()

    solicitante.situacao = Usuario.Situacao.RECUSADO
    solicitante.is_active = False
    solicitante.aprovado_por = request.user
    solicitante.decidido_em = timezone.now()
    solicitante.observacao_decisao = observacao
    solicitante.save()

    messages.success(
        request,
        f"Solicitação de {solicitante.get_full_name() or solicitante.email} recusada."
    )
    return redirect("usuarios_lista")

# ============================================================
# GERENCIAMENTO DE PROJETOS
# ============================================================

@login_required
def projeto_lista(request):
    """
    Lista os projetos cadastrados e a quantidade de imagens ativas
    associadas a cada um.
    """
    from django.db.models import Count, Q
    from django.db.models.functions import Lower

    projetos = (
        Projeto.objects
        .annotate(
            total_imagens=Count(
                "imagens",
                filter=Q(imagens__ativo=True),
            )
        )
        .order_by(Lower("nome"))
    )

    return render(request, "core/projeto_lista.html", {
        "projetos": projetos,
        "pode_gerenciar": _apenas_coordenador(request.user),
    })


@login_required
def projeto_criar(request):
    if not _apenas_coordenador(request.user):
        messages.error(
            request,
            "Você não tem permissão para gerenciar projetos.",
        )
        return redirect("projeto_lista")

    if request.method == "POST":
        nome = request.POST.get("nome", "").strip()
        descricao = request.POST.get("descricao", "").strip()
        ativo = request.POST.get("ativo") == "on"

        if not nome:
            messages.error(request, "Informe o nome do projeto.")
        elif Projeto.objects.filter(nome__iexact=nome).exists():
            messages.error(
                request,
                f"Já existe um projeto chamado '{nome}'.",
            )
        else:
            projeto = Projeto.objects.create(
                nome=nome,
                descricao=descricao,
                ativo=ativo,
            )
            messages.success(
                request,
                f"Projeto '{projeto.nome}' criado.",
            )
            return redirect("projeto_lista")

    return render(request, "core/cadastro_item_form.html", {
        "tipo": "projeto",
        "titulo": "Novo projeto",
        "item": None,
        "url_voltar": "projeto_lista",
    })


@login_required
def projeto_editar(request, pk):
    if not _apenas_coordenador(request.user):
        messages.error(
            request,
            "Você não tem permissão para gerenciar projetos.",
        )
        return redirect("projeto_lista")

    projeto = get_object_or_404(Projeto, pk=pk)

    if request.method == "POST":
        nome = request.POST.get("nome", "").strip()
        descricao = request.POST.get("descricao", "").strip()
        ativo = request.POST.get("ativo") == "on"

        if not nome:
            messages.error(request, "Informe o nome do projeto.")
        elif (
            Projeto.objects
            .filter(nome__iexact=nome)
            .exclude(pk=projeto.pk)
            .exists()
        ):
            messages.error(
                request,
                f"Já existe um projeto chamado '{nome}'.",
            )
        else:
            projeto.nome = nome
            projeto.descricao = descricao
            projeto.ativo = ativo
            projeto.save()

            messages.success(
                request,
                f"Projeto '{projeto.nome}' atualizado.",
            )
            return redirect("projeto_lista")

    return render(request, "core/cadastro_item_form.html", {
        "tipo": "projeto",
        "titulo": "Editar projeto",
        "item": projeto,
        "url_voltar": "projeto_lista",
    })


@login_required
@require_POST
def projeto_toggle_ativo(request, pk):
    if not _apenas_coordenador(request.user):
        messages.error(
            request,
            "Você não tem permissão para gerenciar projetos.",
        )
        return redirect("projeto_lista")

    projeto = get_object_or_404(Projeto, pk=pk)
    projeto.ativo = not projeto.ativo
    projeto.save(update_fields=["ativo", "atualizado_em"])

    estado = "ativado" if projeto.ativo else "inativado"

    messages.success(
        request,
        f"Projeto '{projeto.nome}' {estado}. "
        "Os vínculos existentes com imagens foram preservados.",
    )
    return redirect("projeto_lista")


# ============================================================
# GERENCIAMENTO DE COMPONENTES CURRICULARES
# ============================================================

@login_required
def componente_lista(request):
    """
    Lista os componentes curriculares cadastrados e quantas imagens
    ativas utilizam cada componente.
    """
    from django.db.models import Count, Q
    from django.db.models.functions import Lower

    componentes = (
        ComponenteCurricular.objects
        .annotate(
            total_imagens=Count(
                "imagens",
                filter=Q(imagens__ativo=True),
            )
        )
        .order_by(Lower("nome"))
    )

    return render(request, "core/componente_lista.html", {
        "componentes": componentes,
        "pode_gerenciar": _apenas_coordenador(request.user),
    })


@login_required
def componente_criar(request):
    if not _apenas_coordenador(request.user):
        messages.error(
            request,
            "Você não tem permissão para gerenciar componentes.",
        )
        return redirect("componente_lista")

    if request.method == "POST":
        nome = request.POST.get("nome", "").strip()
        ativo = request.POST.get("ativo") == "on"

        if not nome:
            messages.error(
                request,
                "Informe o nome do componente curricular.",
            )
        elif ComponenteCurricular.objects.filter(nome__iexact=nome).exists():
            messages.error(
                request,
                f"Já existe um componente chamado '{nome}'.",
            )
        else:
            componente = ComponenteCurricular.objects.create(
                nome=nome,
                ativo=ativo,
            )
            messages.success(
                request,
                f"Componente '{componente.nome}' criado.",
            )
            return redirect("componente_lista")

    return render(request, "core/cadastro_item_form.html", {
        "tipo": "componente",
        "titulo": "Novo componente curricular",
        "item": None,
        "url_voltar": "componente_lista",
    })


@login_required
def componente_editar(request, pk):
    if not _apenas_coordenador(request.user):
        messages.error(
            request,
            "Você não tem permissão para gerenciar componentes.",
        )
        return redirect("componente_lista")

    componente = get_object_or_404(ComponenteCurricular, pk=pk)

    if request.method == "POST":
        nome = request.POST.get("nome", "").strip()
        ativo = request.POST.get("ativo") == "on"

        if not nome:
            messages.error(
                request,
                "Informe o nome do componente curricular.",
            )
        elif (
            ComponenteCurricular.objects
            .filter(nome__iexact=nome)
            .exclude(pk=componente.pk)
            .exists()
        ):
            messages.error(
                request,
                f"Já existe um componente chamado '{nome}'.",
            )
        else:
            componente.nome = nome
            componente.ativo = ativo
            componente.save()

            messages.success(
                request,
                f"Componente '{componente.nome}' atualizado.",
            )
            return redirect("componente_lista")

    return render(request, "core/cadastro_item_form.html", {
        "tipo": "componente",
        "titulo": "Editar componente curricular",
        "item": componente,
        "url_voltar": "componente_lista",
    })


@login_required
@require_POST
def componente_toggle_ativo(request, pk):
    if not _apenas_coordenador(request.user):
        messages.error(
            request,
            "Você não tem permissão para gerenciar componentes.",
        )
        return redirect("componente_lista")

    componente = get_object_or_404(ComponenteCurricular, pk=pk)
    componente.ativo = not componente.ativo
    componente.save(update_fields=["ativo", "atualizado_em"])

    estado = "ativado" if componente.ativo else "inativado"

    messages.success(
        request,
        f"Componente '{componente.nome}' {estado}. "
        "Os vínculos existentes com imagens foram preservados.",
    )
    return redirect("componente_lista")


# ============================================================
# GESTÃO DE STATUS DO WORKFLOW (Admin/Coordenador criam e editam,
# todos os perfis podem visualizar a fila completa)
# ============================================================

@login_required
def status_lista(request):
    """
    Lista todos os status do workflow, na ordem da fila. Visível para
    qualquer perfil autenticado (transparência do fluxo); criar, editar,
    reordenar e (des)ativar é restrito a Coordenador/Administrador — a
    view valida isso de novo em cada ação, o template só esconde os
    controles visualmente.
    """
    status_list = StatusWorkflow.objects.all().order_by("ordem")
    return render(request, "core/status_lista.html", {
        "status_list": status_list,
        "pode_gerenciar": _apenas_coordenador(request.user),
    })


def _aplicar_exclusividade_inicial_final(status_obj):
    """
    Garante no máximo um status com is_inicial=True e um com is_final=True
    na fila inteira — desmarca qualquer outro que já tivesse a mesma marca,
    para o motor de workflow (avancar_status) nunca ficar ambíguo sobre
    onde a fila começa ou termina.
    """
    if status_obj.is_inicial:
        StatusWorkflow.objects.filter(is_inicial=True).exclude(pk=status_obj.pk).update(is_inicial=False)
    if status_obj.is_final:
        StatusWorkflow.objects.filter(is_final=True).exclude(pk=status_obj.pk).update(is_final=False)


@login_required
def status_criar(request):
    """Cria um novo status. Somente Coordenador/Administrador."""
    if not _apenas_coordenador(request.user):
        messages.error(request, "Você não tem permissão para gerenciar status.")
        return redirect("status_lista")

    if request.method == "POST":
        from django.db.models import Max

        nome = request.POST.get("nome", "").strip()
        slug = request.POST.get("slug", "").strip()
        descricao_txt = request.POST.get("descricao", "").strip()
        perfil_responsavel = request.POST.get("perfil_responsavel")
        exige_atribuicao = request.POST.get("exige_atribuicao") == "on"
        permite_edicao = request.POST.get("permite_edicao") == "on"
        avanca_ao_abrir = request.POST.get("avanca_ao_abrir") == "on"
        descricao_concluida = request.POST.get("descricao_concluida") == "on"
        conferencia_concluida = request.POST.get("conferencia_concluida") == "on"
        revisao_concluida = request.POST.get("revisao_concluida") == "on"
        is_inicial = request.POST.get("is_inicial") == "on"
        is_final = request.POST.get("is_final") == "on"

        if not nome or not slug:
            messages.error(request, "Nome e identificador interno são obrigatórios.")
            return redirect("status_criar")

        if StatusWorkflow.objects.filter(slug=slug).exists():
            messages.error(request, f"Já existe um status com o identificador '{slug}'.")
            return redirect("status_criar")

        maior_ordem = StatusWorkflow.objects.aggregate(m=Max("ordem"))["m"] or 0

        novo = StatusWorkflow.objects.create(
            nome=nome,
            slug=slug,
            descricao=descricao_txt,
            perfil_responsavel=perfil_responsavel,
            exige_atribuicao=exige_atribuicao,
            is_inicial=is_inicial,
            is_final=is_final,
            ordem=maior_ordem + 1,
            ativo=True,
            permite_edicao=permite_edicao,
            avanca_ao_abrir=avanca_ao_abrir,
            descricao_concluida=descricao_concluida,
            conferencia_concluida=conferencia_concluida,
            revisao_concluida=revisao_concluida,
        )
        _aplicar_exclusividade_inicial_final(novo)

        messages.success(request, f"Status '{novo.nome}' criado.")
        return redirect("status_lista")

    return render(request, "core/status_form.html", {
        "status": None,
        "perfis": StatusWorkflow.PerfilResponsavel.choices,
    })


@login_required
def status_editar(request, pk):
    """Edita um status existente. Somente Coordenador/Administrador."""
    if not _apenas_coordenador(request.user):
        messages.error(request, "Você não tem permissão para gerenciar status.")
        return redirect("status_lista")

    status_obj = get_object_or_404(StatusWorkflow, pk=pk)

    if request.method == "POST":
        nome = request.POST.get("nome", "").strip()
        descricao_txt = request.POST.get("descricao", "").strip()
        perfil_responsavel = request.POST.get("perfil_responsavel")
        exige_atribuicao = request.POST.get("exige_atribuicao") == "on"
        permite_edicao = request.POST.get("permite_edicao") == "on"
        avanca_ao_abrir = request.POST.get("avanca_ao_abrir") == "on"
        descricao_concluida = request.POST.get("descricao_concluida") == "on"
        conferencia_concluida = request.POST.get("conferencia_concluida") == "on"
        revisao_concluida = request.POST.get("revisao_concluida") == "on"
        is_inicial = request.POST.get("is_inicial") == "on"
        is_final = request.POST.get("is_final") == "on"

        if not nome:
            messages.error(request, "O nome é obrigatório.")
            return redirect("status_editar", pk=pk)

        status_obj.nome = nome
        status_obj.descricao = descricao_txt
        status_obj.perfil_responsavel = perfil_responsavel
        status_obj.exige_atribuicao = exige_atribuicao
        status_obj.is_inicial = is_inicial
        status_obj.is_final = is_final
        status_obj.permite_edicao = permite_edicao
        status_obj.avanca_ao_abrir = avanca_ao_abrir
        status_obj.descricao_concluida = descricao_concluida
        status_obj.conferencia_concluida = conferencia_concluida
        status_obj.revisao_concluida = revisao_concluida
        status_obj.save()
        _aplicar_exclusividade_inicial_final(status_obj)

        messages.success(request, f"Status '{status_obj.nome}' atualizado.")
        return redirect("status_lista")

    return render(request, "core/status_form.html", {
        "status": status_obj,
        "perfis": StatusWorkflow.PerfilResponsavel.choices,
    })


@login_required
@require_POST
def status_toggle_ativo(request, pk):
    """
    Ativa/desativa um status. Desativação é bloqueada se existir alguma
    imagem ativa parada nele agora — evita "órfãos" no workflow.
    """
    if not _apenas_coordenador(request.user):
        return JsonResponse({"ok": False, "erro": "Sem permissão."}, status=403)

    status_obj = get_object_or_404(StatusWorkflow, pk=pk)

    if status_obj.ativo:
        em_uso = Imagem.objects.filter(status=status_obj, ativo=True).count()
        if em_uso > 0:
            return JsonResponse({
                "ok": False,
                "erro": f"Não é possível desativar: {em_uso} imagem(ns) está(ão) neste status agora.",
            }, status=400)
        status_obj.ativo = False
    else:
        status_obj.ativo = True

    status_obj.save()
    return JsonResponse({"ok": True, "ativo": status_obj.ativo})


@login_required
@require_POST
def status_reordenar(request):
    """
    Recebe a nova ordem (lista de IDs) via JSON e reatribui o campo
    'ordem' de cada status conforme a posição na lista — usado pelo
    drag-and-drop da tela de gestão.
    """
    if not _apenas_coordenador(request.user):
        return JsonResponse({"ok": False, "erro": "Sem permissão."}, status=403)

    try:
        body = json.loads(request.body)
        ids_em_ordem = body.get("ids", [])
    except (json.JSONDecodeError, AttributeError):
        return JsonResponse({"ok": False, "erro": "JSON inválido."}, status=400)

    if not ids_em_ordem:
        return JsonResponse({"ok": False, "erro": "Lista vazia."}, status=400)

    with transaction.atomic():
        for posicao, status_id in enumerate(ids_em_ordem, start=1):
            StatusWorkflow.objects.filter(pk=status_id).update(ordem=posicao)

    return JsonResponse({"ok": True})


# ============================================================
# MINHA CONTA (autoatendimento — qualquer perfil logado)
# ============================================================

@login_required
def minha_conta(request):
    """
    Tela de autoatendimento: cada usuário edita o próprio nome/sobrenome.
    E-mail, perfil e situação ficam somente leitura — e-mail é o login
    (USERNAME_FIELD) e perfil/situação são geridos pelo Administrador na
    tela de Usuários, não aqui.
    """
    usuario = request.user

    if request.method == "POST":
        first_name = request.POST.get("first_name", "").strip()
        last_name = request.POST.get("last_name", "").strip()

        if not first_name:
            messages.error(request, "O nome é obrigatório.")
            return redirect("minha_conta")

        usuario.first_name = first_name
        usuario.last_name = last_name
        usuario.save(update_fields=["first_name", "last_name"])
        messages.success(request, "Dados atualizados.")
        return redirect("minha_conta")

    return render(request, "core/minha_conta.html", {"usuario": usuario})

# ============================================================
# RELATÓRIOS
# ============================================================
# Exportação no formato compatível com o FotoWeb, gerada só a partir
# de imagens com a descrição finalizada (fluxo completo: descrita,
# conferida e revisada). Enquanto não estiver finalizada, a imagem
# não aparece aqui — o relatório é sempre o "resultado pronto".

import json as _json


def _lang_tag(idioma_codigo):
    """
    Converte o código ISO 639-3 salvo no Trecho (ex: 'por') pra uma tag
    tipo BCP-47 (ex: 'pt-BR'), no mesmo estilo usado pelo FotoWeb.
    Português vira especificamente 'pt-BR'; os demais idiomas usam
    só o código de 2 letras (ex: 'en', 'es').
    """
    try:
        lang = pycountry.languages.get(alpha_3=idioma_codigo)
        alpha_2 = lang.alpha_2 if lang and hasattr(lang, "alpha_2") else None
    except Exception:
        alpha_2 = None

    if not alpha_2:
        return idioma_codigo or "und"
    if alpha_2 == "pt":
        return "pt-BR"
    return alpha_2


def _data_filtro(valor):
    """
    Converte uma data recebida por GET (AAAA-MM-DD) para date.

    Se o valor estiver vazio ou inválido, o filtro é simplesmente ignorado.
    """
    from datetime import date

    valor = (valor or "").strip()
    if not valor:
        return None

    try:
        return date.fromisoformat(valor)
    except ValueError:
        return None


def _imagens_relatorio(request):
    """
    Queryset base do relatório geral.

    Todos os filtros desta função são compartilhados pela tela e pela
    exportação. Assim, o Excel sempre contém exatamente o mesmo conjunto
    encontrado no relatório após aplicar os filtros.
    """
    imagens = (
        Imagem.objects
        .filter(ativo=True)
        .select_related(
            "status",
            "responsavel",
            "lote",
            "projeto",
            "componente_curricular",
            "descricao",
            "descricao__descritor",
            "descricao__revisor",
            "descricao__coordenador",
        )
        .prefetch_related("descricao__trechos")
        .order_by("nome_obra", "retranca")
    )

    retranca_f = request.GET.get("retranca", "").strip()
    if retranca_f:
        imagens = imagens.filter(retranca__icontains=retranca_f)

    projeto_f = request.GET.get("projeto", "").strip()
    if projeto_f:
        if projeto_f == "sem_projeto":
            imagens = imagens.filter(projeto__isnull=True)
        else:
            imagens = imagens.filter(projeto_id=projeto_f)

    obra_f = request.GET.get("obra", "").strip()
    if obra_f:
        imagens = imagens.filter(nome_obra=obra_f)

    lote_f = request.GET.get("lote", "").strip()
    if lote_f:
        if lote_f == "avulsas":
            imagens = imagens.filter(lote__isnull=True)
        else:
            imagens = imagens.filter(lote_id=lote_f)

    componente_f = request.GET.get("componente", "").strip()
    if componente_f:
        imagens = imagens.filter(componente_curricular__nome=componente_f)

    status_f = request.GET.get("status", "").strip()
    if status_f:
        imagens = imagens.filter(status_id=status_f)

    responsavel_f = request.GET.get("responsavel", "").strip()
    if responsavel_f:
        if responsavel_f == "sem_responsavel":
            imagens = imagens.filter(responsavel__isnull=True)
        else:
            imagens = imagens.filter(responsavel_id=responsavel_f)

    pagamento_descritor_f = request.GET.get(
        "pagamento_descritor",
        "",
    ).strip()
    if pagamento_descritor_f:
        imagens = imagens.filter(
            pagamento_descritor=pagamento_descritor_f
        )

    pagamento_revisor_f = request.GET.get(
        "pagamento_revisor",
        "",
    ).strip()
    if pagamento_revisor_f:
        imagens = imagens.filter(
            pagamento_revisor=pagamento_revisor_f
        )

    data_inicio = _data_filtro(request.GET.get("data_inicio"))
    if data_inicio:
        imagens = imagens.filter(criado_em__date__gte=data_inicio)

    data_fim = _data_filtro(request.GET.get("data_fim"))
    if data_fim:
        imagens = imagens.filter(criado_em__date__lte=data_fim)

    return imagens


def _imagens_finalizadas(request):
    """
    Subconjunto das imagens filtradas que já chegou ao status final.

    Mantido para a planilha de compatibilidade com o formato FotoWeb.
    A regra usa a flag is_final do StatusWorkflow, não o nome do status.
    """
    return (
        _imagens_relatorio(request)
        .filter(
            status__is_final=True,
            descricao__finalizado=True,
        )
    )


@login_required
def relatorios_lista(request):
    from django.core.paginator import Paginator
    from django.db.models import Count
    from .models import Lote

    imagens = _imagens_relatorio(request)

    total = imagens.count()
    finalizadas = imagens.filter(status__is_final=True).count()
    em_andamento = total - finalizadas
    sem_lote = imagens.filter(lote__isnull=True).count()

    status_resumo = list(
        imagens
        .values(
            "status_id",
            "status__nome",
            "status__slug",
            "status__ordem",
        )
        .annotate(quantidade=Count("id"))
        .order_by("status__ordem", "status__nome")
    )

    paginator = Paginator(imagens, 50)
    pagina_obj = paginator.get_page(request.GET.get("pagina", 1))

    # Query string compartilhada por exportação e paginação.
    # "pagina" não entra porque o Excel exporta todo o resultado filtrado,
    # não apenas a página atual da tabela.
    filtros_query_dict = request.GET.copy()
    filtros_query_dict.pop("pagina", None)
    filtros_query = filtros_query_dict.urlencode()

    imagens_base = Imagem.objects.filter(ativo=True)

    projetos = (
        Projeto.objects
        .filter(imagens__ativo=True)
        .distinct()
        .order_by("nome")
    )

    componentes = (
        imagens_base
        .filter(componente_curricular__isnull=False)
        .values_list("componente_curricular__nome", flat=True)
        .distinct()
        .order_by("componente_curricular__nome")
    )

    responsaveis = (
        Usuario.objects
        .filter(imagens_responsavel__ativo=True)
        .distinct()
        .order_by("first_name", "last_name", "email")
    )

    status_disponiveis = (
        StatusWorkflow.objects
        .filter(imagens__ativo=True)
        .distinct()
        .order_by("ordem", "nome")
    )

    filtros_ativos = any([
        request.GET.get("retranca"),
        request.GET.get("projeto"),
        request.GET.get("obra"),
        request.GET.get("componente"),
        request.GET.get("lote"),
        request.GET.get("status"),
        request.GET.get("responsavel"),
        request.GET.get("pagamento_descritor"),
        request.GET.get("pagamento_revisor"),
        request.GET.get("data_inicio"),
        request.GET.get("data_fim"),
    ])

    contexto = {
        "pagina_obj": pagina_obj,
        "total": total,
        "finalizadas": finalizadas,
        "em_andamento": em_andamento,
        "sem_lote": sem_lote,
        "status_resumo": status_resumo,

        # Opções dos filtros
        "projetos": projetos,
        "obras": (
            imagens_base
            .exclude(nome_obra="")
            .values_list("nome_obra", flat=True)
            .distinct()
            .order_by("nome_obra")
        ),
        "lotes": Lote.objects.filter(ativo=True).order_by("nome"),
        "componentes": componentes,
        "status_disponiveis": status_disponiveis,
        "responsaveis": responsaveis,
        "pagamentos": Imagem.StatusPagamento.choices,

        # Valores selecionados
        "retranca_selecionada": request.GET.get("retranca", ""),
        "projeto_selecionado": request.GET.get("projeto", ""),
        "obra_selecionada": request.GET.get("obra", ""),
        "componente_selecionado": request.GET.get("componente", ""),
        "lote_selecionado": request.GET.get("lote", ""),
        "status_selecionado": request.GET.get("status", ""),
        "responsavel_selecionado": request.GET.get("responsavel", ""),
        "pagamento_descritor_selecionado": request.GET.get(
            "pagamento_descritor",
            "",
        ),
        "pagamento_revisor_selecionado": request.GET.get(
            "pagamento_revisor",
            "",
        ),
        "data_inicio_selecionada": request.GET.get("data_inicio", ""),
        "data_fim_selecionada": request.GET.get("data_fim", ""),

        "filtros_query": filtros_query,
        "filtros_ativos": filtros_ativos,
    }
    return render(request, "core/relatorios.html", contexto)


@login_required
def relatorios_exportar_fotoweb(request):
    """
    Gera um .xlsx no MESMO modelo do relatório original do FotoWeb.

    A linha base vem do snapshot salvo no momento da importação. Assim,
    campos como keywords, status, img_file, usuario e etapa permanecem
    exatamente como vieram do FotoWeb.

    Somente:
    - descricao
    - descricao_flat

    são substituídas pelas versões atuais existentes no Dito!.

    Os filtros aplicados na tela de Relatórios também são respeitados.
    """
    from django.http import HttpResponse
    from openpyxl import Workbook

    imagens = list(
        _imagens_relatorio(request)
    )

    if not imagens:
        messages.error(
            request,
            "Nenhuma imagem encontrada para exportar.",
        )
        return redirect("relatorios_lista")

    faltando_base = [
        imagem.retranca
        for imagem in imagens
        if not imagem.dados_fotoweb_originais
    ]

    if faltando_base:
        amostra = ", ".join(
            faltando_base[:5]
        )

        complemento = (
            f" Exemplos: {amostra}."
            if amostra
            else ""
        )

        messages.error(
            request,
            (
                f"{len(faltando_base)} imagem(ns) ainda não possuem a "
                "linha original do FotoWeb armazenada."
                + complemento
                + " Reimporte o relatório original dessas imagens; "
                  "as retrancas existentes serão preservadas e somente "
                  "a base FotoWeb será sincronizada."
            ),
        )

        retorno = reverse(
            "relatorios_lista"
        )

        query_string = (
            request.GET.urlencode()
        )

        if query_string:
            retorno = (
                f"{retorno}?{query_string}"
            )

        return redirect(
            retorno
        )

    # Ordem semelhante à origem quando o snapshot possui os metadados.
    imagens.sort(
        key=lambda imagem: (
            str(
                (
                    imagem.dados_fotoweb_originais
                    or {}
                ).get(
                    "__arquivo_origem",
                    "",
                )
                or ""
            ),
            (
                (
                    imagem.dados_fotoweb_originais
                    or {}
                ).get(
                    "__numero_linha",
                    10**12,
                )
                or 10**12
            ),
            imagem.retranca,
        )
    )

    workbook = Workbook()
    worksheet = workbook.active
    worksheet.title = "Sheet1"

    # Cabeçalho idêntico ao modelo original.
    worksheet.append(
        FOTOWEB_COLUNAS_RELATORIO
    )

    for imagem in imagens:
        base = dict(
            imagem.dados_fotoweb_originais
            or {}
        )

        # ----------------------------------------------------
        # A ÚNICA atualização feita no conteúdo da linha:
        # descrição atual do Dito.
        # ----------------------------------------------------
        descricao = getattr(
            imagem,
            "descricao",
            None,
        )

        if descricao:
            trechos = list(
                descricao.trechos
                .filter(
                    ativo=True
                )
                .order_by(
                    "ordem"
                )
            )

            if trechos:
                base["descricao"] = (
                    _json.dumps(
                        [
                            {
                                "lang": _lang_tag(
                                    trecho.idioma_codigo
                                ),
                                "text": trecho.texto,
                            }
                            for trecho in trechos
                        ],
                        ensure_ascii=False,
                    )
                )

                base["descricao_flat"] = (
                    " ".join(
                        trecho.texto
                        for trecho in trechos
                    ).strip()
                )

        worksheet.append(
            [
                base.get(coluna)
                for coluna in FOTOWEB_COLUNAS_RELATORIO
            ]
        )

    resposta = HttpResponse(
        content_type=(
            "application/vnd.openxmlformats-officedocument."
            "spreadsheetml.sheet"
        )
    )

    resposta["Content-Disposition"] = (
        'attachment; filename="relatorio_fotoweb_atualizado.xlsx"'
    )

    workbook.save(
        resposta
    )

    return resposta


@login_required
def relatorios_exportar(request):
    """
    Exporta um arquivo .xlsx com duas abas.

    A função usa _imagens_relatorio(request), portanto TODOS os filtros
    escolhidos na tela também são respeitados no arquivo exportado.

    1. Geral
       Todas as imagens que respeitam os filtros atuais.

    2. FotoWeb - Finalizados
       Apenas as imagens filtradas que já chegaram ao status final.
    """
    from django.http import HttpResponse
    from django.utils import timezone
    from openpyxl import Workbook

    def _valor_data_excel(valor, incluir_hora=False):
        """
        Converte datas para texto antes de enviá-las ao openpyxl.

        Django trabalha com DateTimeField timezone-aware quando USE_TZ=True,
        enquanto o Excel/openpyxl não aceita datetimes com fuso horário.
        """
        if not valor:
            return ""

        if hasattr(valor, "tzinfo") and valor.tzinfo is not None:
            if timezone.is_aware(valor):
                valor = timezone.localtime(valor)

        if incluir_hora:
            return valor.strftime("%d/%m/%Y %H:%M")

        return valor.strftime("%d/%m/%Y")

    imagens = list(_imagens_relatorio(request))
    imagens_finalizadas = [
        imagem
        for imagem in imagens
        if imagem.status
        and imagem.status.is_final
        and getattr(getattr(imagem, "descricao", None), "finalizado", False)
    ]

    wb = Workbook()

    # ========================================================
    # ABA 1 — RELATÓRIO GERAL
    # ========================================================
    aba_geral = wb.active
    aba_geral.title = "Geral"

    colunas_geral = [
        "Retranca",
        "Obra",
        "Componente curricular",
        "Volume/Ano/Módulo",
        "Capítulo/Unidade",
        "Etapa",
        "Lote",
        "Data prevista do lote",
        "Data efetiva do lote",
        "Status",
        "Responsável atual",
        "Descritor",
        "Revisor",
        "Coordenador",
        "Pagamento descritor",
        "Pagamento revisor",
        "Prazo da imagem",
        "Cadastrada em",
        "Atualizada em",
        "Descrição completa",
    ]
    aba_geral.append(colunas_geral)

    for imagem in imagens:
        descricao = getattr(imagem, "descricao", None)

        if descricao:
            trechos = list(
                descricao.trechos
                .filter(ativo=True)
                .order_by("ordem")
            )
            descricao_flat = " ".join(t.texto for t in trechos)
        else:
            descricao_flat = ""

        lote = imagem.lote

        responsavel = (
            imagem.responsavel.get_full_name()
            or imagem.responsavel.email
            if imagem.responsavel
            else ""
        )
        descritor = (
            descricao.descritor.get_full_name()
            or descricao.descritor.email
            if descricao and descricao.descritor
            else ""
        )
        revisor = (
            descricao.revisor.get_full_name()
            or descricao.revisor.email
            if descricao and descricao.revisor
            else ""
        )
        coordenador = (
            descricao.coordenador.get_full_name()
            or descricao.coordenador.email
            if descricao and descricao.coordenador
            else ""
        )

        aba_geral.append([
            imagem.retranca,
            imagem.nome_obra,
            imagem.componente_curricular.nome if imagem.componente_curricular else "",
            imagem.volume_ano_modulo,
            imagem.capitulo_unidade,
            imagem.get_etapa_display(),
            lote.nome if lote else "",
            _valor_data_excel(
                lote.data_prevista if lote else None
            ),
            _valor_data_excel(
                lote.data_efetiva if lote else None
            ),
            imagem.status.nome if imagem.status else "",
            responsavel,
            descritor,
            revisor,
            coordenador,
            imagem.get_pagamento_descritor_display(),
            imagem.get_pagamento_revisor_display(),
            _valor_data_excel(imagem.prazo),
            _valor_data_excel(
                imagem.criado_em,
                incluir_hora=True,
            ),
            _valor_data_excel(
                imagem.atualizado_em,
                incluir_hora=True,
            ),
            descricao_flat or None,
        ])

    larguras_geral = [
        28, 22, 20, 18, 18, 12, 18, 18, 18, 22,
        24, 24, 24, 24, 20, 20, 16, 20, 20, 50,
    ]
    for i, largura in enumerate(larguras_geral, start=1):
        aba_geral.column_dimensions[
            aba_geral.cell(row=1, column=i).column_letter
        ].width = largura

    # ========================================================
    # ABA 2 — FORMATO FOTOWEB (SÓ FINALIZADOS)
    # ========================================================
    aba_fotoweb = wb.create_sheet("FotoWeb - Finalizados")

    colunas_fotoweb = [
        "obra",
        "componente",
        "volume",
        "capitulo",
        "keywords",
        "status",
        "retranca",
        "img_file",
        "descricao",
        "usuario",
        "etapa",
        "retranca_lower",
        "descricao_flat",
    ]
    aba_fotoweb.append(colunas_fotoweb)

    for imagem in imagens_finalizadas:
        descricao = imagem.descricao
        trechos = list(
            descricao.trechos
            .filter(ativo=True)
            .order_by("ordem")
        )

        trechos_json = [
            {
                "lang": _lang_tag(t.idioma_codigo),
                "text": t.texto,
            }
            for t in trechos
        ]
        descricao_flat = " ".join(t.texto for t in trechos)

        usuario = descricao.descritor or imagem.responsavel
        usuario_login = usuario.username if usuario else ""

        aba_fotoweb.append([
            imagem.nome_obra,
            imagem.componente_curricular.nome if imagem.componente_curricular else "",
            imagem.volume_ano_modulo,
            imagem.capitulo_unidade,
            "",
            imagem.status.nome if imagem.status else "",
            imagem.retranca,
            imagem.url_fotoweb,
            _json.dumps(
                trechos_json,
                ensure_ascii=False,
            ) if trechos_json else None,
            usuario_login,
            f"Etapa: {imagem.etapa}",
            imagem.retranca.lower(),
            descricao_flat or None,
        ])

    for i, largura in enumerate(
        [16, 14, 8, 10, 10, 14, 24, 40, 40, 14, 12, 24, 40],
        start=1,
    ):
        aba_fotoweb.column_dimensions[
            aba_fotoweb.cell(row=1, column=i).column_letter
        ].width = largura

    resposta = HttpResponse(
        content_type=(
            "application/vnd.openxmlformats-officedocument."
            "spreadsheetml.sheet"
        )
    )
    resposta["Content-Disposition"] = (
        'attachment; filename="relatorio_geral_dito.xlsx"'
    )
    wb.save(resposta)
    return resposta


# ============================================================
# BUSCAR RETRANCA
# ============================================================

@login_required
def buscar_retranca(request):
    from .models import filtro_autoria_imagem

    termo = request.GET.get("q", "").strip()
    resultados = []

    if termo:
        imagens = Imagem.objects.filter(ativo=True).select_related("status", "lote")

        # Descritor/revisor só vê o que é responsabilidade dele; coordenação/admin vê tudo.
        if not _apenas_coordenador(request.user):
            imagens = imagens.filter(filtro_autoria_imagem(request.user)).distinct()

        resultados = imagens.filter(retranca__icontains=termo).order_by("nome_obra", "retranca")[:100]

    return render(request, "core/buscar_retranca.html", {
        "termo": termo,
        "resultados": resultados,
    })


# ============================================================
# HISTÓRICO
# ============================================================

@login_required
def historico_lista(request):
    from django.core.paginator import Paginator
    from .models import HistoricoItem

    if not _apenas_coordenador(request.user):
        messages.error(request, "Você não tem acesso ao histórico geral.")
        return redirect("dashboard")

    itens = HistoricoItem.objects.select_related(
        "imagem", "usuario", "status_anterior", "novo_status"
    ).order_by("-criado_em")

    busca = request.GET.get("busca", "").strip()
    if busca:
        itens = itens.filter(imagem__retranca__icontains=busca)

    tipo_acao_f = request.GET.get("tipo_acao", "").strip()
    if tipo_acao_f:
        itens = itens.filter(tipo_acao=tipo_acao_f)

    usuario_f = request.GET.get("usuario", "").strip()
    if usuario_f:
        itens = itens.filter(usuario_id=usuario_f)

    paginator = Paginator(itens, 40)
    pagina_obj = paginator.get_page(request.GET.get("pagina", 1))

    return render(request, "core/historico_lista.html", {
        "pagina_obj": pagina_obj,
        "busca": busca,
        "tipo_acao_f": tipo_acao_f,
        "usuario_f": usuario_f,
        "tipos_acao": HistoricoItem.TipoAcao.choices,
        "usuarios": Usuario.objects.filter(acoes__isnull=False).distinct().order_by("first_name"),
    })

@login_required
@require_POST
def lote_alterar_status(request, lote_id):
    """
    Altera manualmente o status de TODAS as imagens de um lote.

    Atalho operacional para coordenação/administração — não passa pelas
    regras normais do motor de workflow. Além do status, mantém o responsável
    coerente com o perfil configurado no status de destino:

    - se o destino pertence à Coordenação, quem executou a ação assume;
    - se o responsável atual já pertence ao perfil do destino, ele é mantido;
    - caso contrário, o responsável é limpo para evitar uma atribuição
      incompatível com a nova etapa.

    Cada imagem gera registro no histórico.
    """
    from .models import HistoricoItem, Lote

    if not _apenas_coordenador(request.user):
        messages.error(
            request,
            "Você não tem permissão para alterar o status de um lote.",
        )
        return redirect("lotes_lista")

    lote = get_object_or_404(Lote, pk=lote_id)
    destino_id = request.POST.get("status_id")

    if not destino_id:
        messages.error(request, "Selecione o status de destino.")
        return redirect(request.POST.get("next", "lotes_lista"))

    try:
        novo_status = StatusWorkflow.objects.get(
            pk=destino_id,
            ativo=True,
        )
    except StatusWorkflow.DoesNotExist:
        messages.error(request, "Status inválido.")
        return redirect(request.POST.get("next", "lotes_lista"))

    imagens = list(
        Imagem.objects
        .filter(lote=lote, ativo=True)
        .select_related("status", "responsavel", "descricao")
        .exclude(status=novo_status)
    )

    if not imagens:
        messages.info(
            request,
            f"Nenhuma imagem do lote '{lote.nome}' precisou ser alterada.",
        )
        return redirect(request.POST.get("next", "lotes_lista"))

    perfil_operacional = _perfil_operacional(request.user)

    with transaction.atomic():
        for img in imagens:
            status_anterior = img.status
            responsavel_anterior = img.responsavel

            # ----------------------------------------------------
            # Responsável coerente com o status de destino
            # ----------------------------------------------------
            if novo_status.perfil_responsavel == perfil_operacional:
                novo_responsavel = request.user
            elif (
                responsavel_anterior
                and responsavel_anterior.tipo
                == novo_status.perfil_responsavel
            ):
                novo_responsavel = responsavel_anterior
            else:
                novo_responsavel = None

            img.status = novo_status
            img.responsavel = novo_responsavel
            img.save(
                update_fields=[
                    "status",
                    "responsavel",
                    "atualizado_em",
                ]
            )

            descricao = getattr(img, "descricao", None)

            if descricao:
                campos_descricao = []

                # Reabre a etapa apenas quando o status representa uma etapa
                # operacional daquele perfil.
                if (
                    novo_status.perfil_responsavel
                    == Usuario.Tipo.DESCRITOR
                    and (
                        novo_status.permite_edicao
                        or novo_status.avanca_ao_abrir
                    )
                ):
                    if descricao.descritor_bloqueado:
                        descricao.descritor_bloqueado = False
                        campos_descricao.append("descritor_bloqueado")

                    if (
                        novo_responsavel
                        and descricao.descritor_id != novo_responsavel.id
                    ):
                        descricao.descritor = novo_responsavel
                        campos_descricao.append("descritor")

                elif (
                    novo_status.perfil_responsavel
                    == Usuario.Tipo.REVISOR
                    and (
                        novo_status.permite_edicao
                        or novo_status.avanca_ao_abrir
                    )
                ):
                    if descricao.revisor_bloqueado:
                        descricao.revisor_bloqueado = False
                        campos_descricao.append("revisor_bloqueado")

                    if (
                        novo_responsavel
                        and descricao.revisor_id != novo_responsavel.id
                    ):
                        descricao.revisor = novo_responsavel
                        campos_descricao.append("revisor")

                elif (
                    novo_status.perfil_responsavel
                    == Usuario.Tipo.COORDENADOR
                    and novo_responsavel
                    and descricao.coordenador_id != novo_responsavel.id
                ):
                    descricao.coordenador = novo_responsavel
                    campos_descricao.append("coordenador")

                finalizado = bool(novo_status.is_final)
                if descricao.finalizado != finalizado:
                    descricao.finalizado = finalizado
                    campos_descricao.append("finalizado")

                if campos_descricao:
                    descricao.save(update_fields=campos_descricao)

            responsavel_antes_txt = (
                str(responsavel_anterior)
                if responsavel_anterior
                else "sem responsável"
            )
            responsavel_depois_txt = (
                str(novo_responsavel)
                if novo_responsavel
                else "sem responsável"
            )

            HistoricoItem.objects.create(
                imagem=img,
                descricao=descricao,
                usuario=request.user,
                tipo_acao=HistoricoItem.TipoAcao.STATUS_ALTERADO,
                status_anterior=status_anterior,
                novo_status=novo_status,
                observacao=(
                    f"Alteração manual de status do lote '{lote.nome}' "
                    f"por {request.user.get_full_name() or request.user.email}. "
                    f"Responsável: {responsavel_antes_txt} → "
                    f"{responsavel_depois_txt}."
                ),
            )

    lote.sincronizar_data_efetiva()

    messages.success(
        request,
        (
            f"{len(imagens)} imagem(ns) do lote '{lote.nome}' "
            f"alterada(s) para '{novo_status.nome}', com os responsáveis "
            "ajustados conforme o perfil do status."
        ),
    )
    return redirect(request.POST.get("next", "lotes_lista"))

