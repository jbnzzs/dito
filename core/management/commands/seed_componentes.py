
from django.core.management.base import BaseCommand
from django.db import transaction

from core.models import ComponenteCurricular


COMPONENTES = [
    "Alfabetização",
    "Arte",
    "Ciências",
    "Educação digital",
    "Educação física",
    "Educação infantil",
    "Geografia",
    "História",
    "Língua Espanhola",
    "Língua Inglesa",
    "Língua Portuguesa",
    "Matemática",
    "Produção de texto",
]


class Command(BaseCommand):
    help = "Cadastra componentes iniciais sem duplicar existentes."

    def add_arguments(self, parser):
        parser.add_argument(
            "--aplicar",
            action="store_true",
            help="Autoriza a criação dos componentes.",
        )

    def handle(self, *args, **options):
        if not options["aplicar"]:
            self.stdout.write(
                "Seed preparado. Use --aplicar para executar."
            )
            return

        criados = 0
        existentes = 0

        with transaction.atomic():
            for nome in COMPONENTES:
                if ComponenteCurricular.objects.filter(
                    nome__iexact=nome
                ).exists():
                    existentes += 1
                    continue

                _, created = ComponenteCurricular.objects.get_or_create(
                    nome=nome
                )

                if created:
                    criados += 1
                else:
                    existentes += 1

        self.stdout.write(
            self.style.SUCCESS(
                f"{criados} criados, {existentes} existentes."
            )
        )
