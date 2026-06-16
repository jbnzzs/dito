from django.core.management.base import BaseCommand
from core.models import StatusWorkflow


STATUS_INICIAIS = [
    {"ordem": 1, "nome": "Liberado para descrição",  "slug": "liberado-descricao"},
    {"ordem": 2, "nome": "Descrevendo",               "slug": "descrevendo"},
    {"ordem": 3, "nome": "Descrito",                  "slug": "descrito"},
    {"ordem": 4, "nome": "Liberado para conferência", "slug": "liberado-conferencia"},
    {"ordem": 5, "nome": "Em conferência",            "slug": "em-conferencia"},
    {"ordem": 6, "nome": "Conferido",                 "slug": "conferido"},
    {"ordem": 7, "nome": "Revisando",                 "slug": "revisando"},
    {"ordem": 8, "nome": "Revisado",                  "slug": "revisado"},
    {"ordem": 9, "nome": "Finalizado",                "slug": "finalizado"},
]


class Command(BaseCommand):
    help = "Popula os status iniciais do workflow editorial do Dito!"

    def handle(self, *args, **kwargs):
        criados = 0
        atualizados = 0

        for dados in STATUS_INICIAIS:
            obj, created = StatusWorkflow.objects.update_or_create(
                slug=dados["slug"],
                defaults={
                    "nome": dados["nome"],
                    "ordem": dados["ordem"],
                    "ativo": True,
                },
            )
            if created:
                criados += 1
                self.stdout.write(self.style.SUCCESS(f"  ✓ Criado: {obj}"))
            else:
                atualizados += 1
                self.stdout.write(f"  ~ Já existe: {obj}")

        self.stdout.write(
            self.style.SUCCESS(
                f"\nConcluído: {criados} criados, {atualizados} já existiam."
            )
        )