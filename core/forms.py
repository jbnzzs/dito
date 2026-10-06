from django import forms

from .models import ComponenteCurricular, Imagem, Projeto

class ImagemForm(forms.ModelForm):
    """
    Formulário de criação e edição de uma Imagem.

    Projeto e Componente Curricular passam a vir de cadastros gerenciáveis.
    Itens inativos não aparecem para novas associações, mas continuam
    disponíveis quando já estão vinculados à imagem que está sendo editada.
    """

    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)

        projetos = Projeto.objects.filter(ativo=True)
        componentes = ComponenteCurricular.objects.filter(ativo=True)

        if self.instance and self.instance.pk:
            if self.instance.projeto_id:
                projetos = (
                    projetos
                    | Projeto.objects.filter(pk=self.instance.projeto_id)
                ).distinct()

            if self.instance.componente_curricular_id:
                componentes = (
                    componentes
                    | ComponenteCurricular.objects.filter(
                        pk=self.instance.componente_curricular_id
                    )
                ).distinct()

        self.fields["projeto"].queryset = projetos.order_by("nome")
        self.fields["componente_curricular"].queryset = componentes.order_by("nome")

        self.fields["projeto"].empty_label = "Sem projeto"
        self.fields["componente_curricular"].empty_label = "Sem componente"

    class Meta:
        model = Imagem
        fields = [
            "retranca",
            "projeto",
            "nome_obra",
            "volume_ano_modulo",
            "componente_curricular",
            "capitulo_unidade",
            "etapa",
            "status",
            "responsavel",
            "prazo",
            "url_fotoweb",
            "url_pdf",
        ]
        widgets = {
            "retranca": forms.TextInput(attrs={"class": "form-control"}),
            "projeto": forms.Select(attrs={"class": "form-select"}),
            "nome_obra": forms.TextInput(attrs={"class": "form-control"}),
            "volume_ano_modulo": forms.TextInput(attrs={"class": "form-control"}),
            "componente_curricular": forms.Select(attrs={"class": "form-select"}),
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


from django.contrib.auth.password_validation import validate_password
from django.core.exceptions import ValidationError

from .models import Usuario


class SolicitacaoAcessoForm(forms.ModelForm):
    """
    Formulário público de solicitação de acesso ao Dito!.
    O usuário é criado como PENDENTE e inativo, aguardando aprovação
    do administrador, que define o perfil no momento da aprovação.
    """
    senha = forms.CharField(
        label="Senha",
        widget=forms.PasswordInput(attrs={"placeholder": "Mínimo de 8 caracteres"}),
    )
    senha_confirmacao = forms.CharField(
        label="Confirmação de senha",
        widget=forms.PasswordInput(attrs={"placeholder": "Digite a senha novamente"}),
    )

    class Meta:
        model = Usuario
        fields = ["first_name", "last_name", "email"]
        labels = {
            "first_name": "Nome",
            "last_name": "Sobrenome",
            "email": "E-mail",
        }
        widgets = {
            "first_name": forms.TextInput(attrs={"placeholder": "Seu nome"}),
            "last_name": forms.TextInput(attrs={"placeholder": "Seu sobrenome"}),
            "email": forms.EmailInput(attrs={"placeholder": "seu.email@scriba.com.br"}),
        }

    def clean_email(self):
        email = self.cleaned_data["email"].strip().lower()
        if Usuario.objects.filter(email__iexact=email).exists():
            raise ValidationError(
                "Já existe uma conta ou solicitação com este e-mail. "
                "Se você já solicitou acesso, aguarde a aprovação do administrador."
            )
        return email

    def clean_senha(self):
        senha = self.cleaned_data["senha"]
        validate_password(senha)
        return senha

    def clean(self):
        dados = super().clean()
        senha = dados.get("senha")
        confirmacao = dados.get("senha_confirmacao")
        if senha and confirmacao and senha != confirmacao:
            self.add_error("senha_confirmacao", "As senhas não coincidem.")
        return dados

    def save(self, commit=True):
        usuario = super().save(commit=False)
        usuario.set_password(self.cleaned_data["senha"])
        usuario.situacao = Usuario.Situacao.PENDENTE
        usuario.is_active = False          # não consegue logar até ser aprovado
        usuario.username = Usuario.objects._gerar_username(usuario.email)
        if commit:
            usuario.save()
        return usuario