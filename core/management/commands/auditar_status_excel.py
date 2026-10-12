"""Audita valores de status de um Excel FotoWeb sem utilizar o banco."""
from django.core.management.base import BaseCommand, CommandError
from core.services.importacao_service import auditar_arquivo_excel


class Command(BaseCommand):
    help = "Conta os status encontrados em um relatório .xlsx, sem importar imagens."
    requires_system_checks = []

    def add_arguments(self, parser):
        parser.add_argument("arquivo", help="Relatório FotoWeb (.xlsx)")
        parser.add_argument("--amostras", type=int, default=3,
                            help="Quantidade de retrancas de exemplo por status (0 a 10).")

    def handle(self, *args, **options):
        amostras = options["amostras"]
        if not 0 <= amostras <= 10:
            raise CommandError("--amostras deve estar entre 0 e 10.")
        try:
            resultado = auditar_arquivo_excel(options["arquivo"], max_amostras=amostras)
        except (OSError, ValueError, RuntimeError, ImportError) as exc:
            raise CommandError(str(exc)) from exc

        self.stdout.write("=== AUDITORIA DE STATUS — SOMENTE LEITURA ===")
        self.stdout.write(f"Linhas do Excel: {resultado['total_linhas']}")
        self.stdout.write(f"Linhas sem retranca: {resultado['sem_retranca']}")
        self.stdout.write(f"Linhas com retranca: {resultado['com_retranca']}")
        self.stdout.write(f"Linhas com retranca e status vazio: {resultado['sem_status']}")
        self.stdout.write("\nValores de STATUS encontrados:")
        for status, quantidade in resultado["status"].items():
            self.stdout.write(f"  {status}: {quantidade}")
            if amostras:
                exemplos = ", ".join(resultado["amostras"].get(status, []))
                self.stdout.write(f"    Exemplo(s): {exemplos}")
        self.stdout.write("\n[OK] Somente leitura. Nenhuma imagem foi importada ou modificada.")
        self.stdout.write("[AVISO] Esta auditoria não verifica os status cadastrados no banco.")
