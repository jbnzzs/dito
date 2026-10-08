from django.db import migrations, models


class Migration(migrations.Migration):

    dependencies = [
        ("core", "0023_imagem_dados_fotoweb_originais"),
    ]

    operations = [
        migrations.AddField(
            model_name="projeto",
            name="acervo_fotoweb",
            field=models.CharField(
                blank=True,
                help_text=(
                    "Nome do acervo usado para gerar automaticamente "
                    "os links das imagens no FotoWeb."
                ),
                max_length=180,
                verbose_name="Acervo FotoWeb",
            ),
        ),
    ]
