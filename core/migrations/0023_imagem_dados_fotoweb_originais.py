from django.db import migrations, models


class Migration(migrations.Migration):

    dependencies = [
        ("core", "0022_componentecurricular_projeto_and_more"),
    ]

    operations = [
        migrations.AddField(
            model_name="imagem",
            name="dados_fotoweb_originais",
            field=models.JSONField(
                blank=True,
                default=dict,
                help_text=(
                    "Snapshot da linha original importada do relatório FotoWeb. "
                    "Usado para gerar um novo arquivo no mesmo formato, alterando "
                    "somente os campos de descrição."
                ),
                verbose_name="Dados originais do FotoWeb",
            ),
        ),
    ]
