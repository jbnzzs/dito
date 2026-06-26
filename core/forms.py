from django import forms

from .models import Imagem

class ImagemForm(forms.ModelForm):
    """
    Formulário de criação e edição de uma Imagem (registro editorial).

    As informações técnicas (nome/caminho/tamanho do arquivo, dimensões,
    resolução, tamanho físico) NÃO entram aqui: vêm da importação de Excel.
    Os campos 'cadastrado_por', 'ativo', 'criado_em' e 'atualizado_em'
    também ficam de fora — são controlados pelo sistema.
    """

    class Meta:
        model = Imagem
        fields = [
            # Metadados editoriais
            "retranca",
            "nome_obra",
            "volume_ano_modulo",
            "componente_curricular",
            "capitulo_unidade",
            "etapa",
            # Workflow
            "status",
            "responsavel",
            "prazo",
            # FotoWeb
            "url_fotoweb",
            "url_pdf",
        ]
        widgets = {
            "retranca": forms.TextInput(attrs={"class": "form-control"}),
            "nome_obra": forms.TextInput(attrs={"class": "form-control"}),
            "volume_ano_modulo": forms.TextInput(attrs={"class": "form-control"}),
            "componente_curricular": forms.TextInput(attrs={"class": "form-control"}),
            "capitulo_unidade": forms.TextInput(attrs={"class": "form-control"}),
            "etapa": forms.Select(attrs={"class": "form-select"}),
            "status": forms.Select(attrs={"class": "form-select"}),
            "responsavel": forms.Select(attrs={"class": "form-select"}),
            "prazo": forms.DateInput(
                attrs={"class": "form-control", "type": "date"},
                format="%Y-%m-%d",
            ),
            "url_fotoweb": forms.URLInput(attrs={
                "class": "form-control",
                "placeholder": "http://fotoweb.ensinolivre.com.br:9090/fotoweb/archives/.../arquivo.png.info",
            }),
            "url_pdf": forms.URLInput(attrs={
                "class": "form-control",
                "placeholder": "https://ensinolivre-my.sharepoint.com/",
            }),
        }