from django.db import migrations, models
import django.db.models.deletion


def migrar_componentes_existentes(apps, schema_editor):
    """
    Converte os valores de texto já existentes em
    Imagem.componente_curricular para o novo cadastro
    ComponenteCurricular.

    O campo antigo continua existindo durante esta etapa.
    O novo vínculo temporário se chama componente_ref.
    """
    ComponenteCurricular = apps.get_model("core", "ComponenteCurricular")
    Imagem = apps.get_model("core", "Imagem")

    cache = {}

    for imagem in Imagem.objects.all().iterator(chunk_size=500):
        nome = (imagem.componente_curricular or "").strip()

        if not nome:
            continue

        chave = nome.casefold()
        componente_id = cache.get(chave)

        if componente_id is None:
            componente, _ = ComponenteCurricular.objects.get_or_create(
                nome=nome,
                defaults={"ativo": True},
            )
            componente_id = componente.pk
            cache[chave] = componente_id

        Imagem.objects.filter(pk=imagem.pk).update(
            componente_ref_id=componente_id,
        )


class Migration(migrations.Migration):

    dependencies = [
        ("core", "0021_remove_lote_prazo"),
    ]

    operations = [
        migrations.CreateModel(
            name="Projeto",
            fields=[
                (
                    "id",
                    models.BigAutoField(
                        auto_created=True,
                        primary_key=True,
                        serialize=False,
                        verbose_name="ID",
                    ),
                ),
                (
                    "nome",
                    models.CharField(
                        max_length=200,
                        unique=True,
                        verbose_name="Nome do projeto",
                    ),
                ),
                (
                    "descricao",
                    models.CharField(
                        blank=True,
                        max_length=255,
                        verbose_name="Descrição",
                    ),
                ),
                (
                    "ativo",
                    models.BooleanField(
                        default=True,
                        verbose_name="Ativo",
                    ),
                ),
                (
                    "criado_em",
                    models.DateTimeField(
                        auto_now_add=True,
                        verbose_name="Criado em",
                    ),
                ),
                (
                    "atualizado_em",
                    models.DateTimeField(
                        auto_now=True,
                        verbose_name="Atualizado em",
                    ),
                ),
            ],
            options={
                "verbose_name": "Projeto",
                "verbose_name_plural": "Projetos",
                "ordering": ["nome"],
            },
        ),

        migrations.CreateModel(
            name="ComponenteCurricular",
            fields=[
                (
                    "id",
                    models.BigAutoField(
                        auto_created=True,
                        primary_key=True,
                        serialize=False,
                        verbose_name="ID",
                    ),
                ),
                (
                    "nome",
                    models.CharField(
                        max_length=100,
                        unique=True,
                        verbose_name="Componente curricular",
                    ),
                ),
                (
                    "ativo",
                    models.BooleanField(
                        default=True,
                        verbose_name="Ativo",
                    ),
                ),
                (
                    "criado_em",
                    models.DateTimeField(
                        auto_now_add=True,
                        verbose_name="Criado em",
                    ),
                ),
                (
                    "atualizado_em",
                    models.DateTimeField(
                        auto_now=True,
                        verbose_name="Atualizado em",
                    ),
                ),
            ],
            options={
                "verbose_name": "Componente curricular",
                "verbose_name_plural": "Componentes curriculares",
                "ordering": ["nome"],
            },
        ),

        migrations.AddField(
            model_name="imagem",
            name="projeto",
            field=models.ForeignKey(
                blank=True,
                null=True,
                on_delete=django.db.models.deletion.PROTECT,
                related_name="imagens",
                to="core.projeto",
                verbose_name="Projeto",
                help_text=(
                    "Agrupador maior que pode reunir imagens "
                    "de diferentes obras/editoras."
                ),
            ),
        ),

        # Campo temporário: evita tentar converter textos como
        # "ARTE" diretamente para bigint.
        migrations.AddField(
            model_name="imagem",
            name="componente_ref",
            field=models.ForeignKey(
                blank=True,
                null=True,
                on_delete=django.db.models.deletion.PROTECT,
                related_name="imagens",
                to="core.componentecurricular",
                verbose_name="Componente curricular",
            ),
        ),

        # Copia os textos antigos para o novo cadastro/FK.
        migrations.RunPython(
            migrar_componentes_existentes,
            migrations.RunPython.noop,
        ),

        # Só depois da cópia removemos o campo texto antigo.
        migrations.RemoveField(
            model_name="imagem",
            name="componente_curricular",
        ),

        # O novo FK assume o nome definitivo esperado pelo model.
        migrations.RenameField(
            model_name="imagem",
            old_name="componente_ref",
            new_name="componente_curricular",
        ),
    ]
