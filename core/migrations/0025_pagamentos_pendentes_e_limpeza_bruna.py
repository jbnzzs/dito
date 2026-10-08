from django.db import migrations, models


# AJUSTES_RAPIDOS_PAGAMENTO_PENDENTE


def aplicar_ajustes(apps, schema_editor):
    Imagem = apps.get_model("core", "Imagem")
    Usuario = apps.get_model("core", "Usuario")

    Imagem.objects.all().update(
        pagamento_descritor="pendente",
        pagamento_revisor="pendente",
    )

    admins_bruna = Usuario.objects.filter(
        first_name__iexact="Bruna",
        last_name__iexact="Germano",
        situacao="aprovado",
        tipo="administrador",
    )

    if admins_bruna.exists():
        Usuario.objects.filter(
            first_name__iexact="Bruna",
            last_name__iexact="Germano",
            situacao="pendente",
        ).delete()


def desfazer_ajustes(apps, schema_editor):
    Imagem = apps.get_model("core", "Imagem")
    Imagem.objects.all().update(
        pagamento_descritor="contabilizado",
        pagamento_revisor="contabilizado",
    )


class Migration(migrations.Migration):

    dependencies = [
        ("core", "0024_projeto_acervo_fotoweb"),
    ]

    operations = [
        migrations.AlterField(
            model_name="imagem",
            name="pagamento_descritor",
            field=models.CharField(
                choices=[
                    ("pendente", "Pendente"),
                    ("contabilizado", "Contabilizado"),
                    ("pago", "Pago"),
                ],
                default="pendente",
                max_length=20,
                verbose_name="Pagamento — Descritor",
            ),
        ),
        migrations.AlterField(
            model_name="imagem",
            name="pagamento_revisor",
            field=models.CharField(
                choices=[
                    ("pendente", "Pendente"),
                    ("contabilizado", "Contabilizado"),
                    ("pago", "Pago"),
                ],
                default="pendente",
                max_length=20,
                verbose_name="Pagamento — Revisor",
            ),
        ),
        migrations.RunPython(
            aplicar_ajustes,
            desfazer_ajustes,
        ),
    ]
