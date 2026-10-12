"""Testes de auditoria dos status (sem acesso ao banco)."""
import tempfile
import unittest
from pathlib import Path
from openpyxl import Workbook
from core.services.importacao_service import (
    normalizar_cabecalho, normalizar_status, resumir_status_linhas,
    auditar_arquivo_excel,
)


class ImportacaoStatusTests(unittest.TestCase):
    def test_normalizacao(self):
        self.assertEqual(normalizar_cabecalho(" STÁTUS "), "status")
        self.assertEqual(normalizar_status("  Conferídas "), "CONFERIDAS")
        self.assertEqual(normalizar_status(None), "<VAZIO>")

    def test_contagem_sem_retranca(self):
        resumo = resumir_status_linhas(
            ["Retranca", "Status"],
            [("r1", "Descritas"), ("r2", "Conferidas"),
             ("r3", " descritas "), ("", "Conferidas"), ("r4", None)],
        )
        self.assertEqual(resumo["total_linhas"], 5)
        self.assertEqual(resumo["sem_retranca"], 1)
        self.assertEqual(resumo["status"]["DESCRITAS"], 2)
        self.assertEqual(resumo["status"]["CONFERIDAS"], 1)
        self.assertEqual(resumo["sem_status"], 1)

    def test_colunas_necessarias(self):
        with self.assertRaises(ValueError):
            resumir_status_linhas(["retranca", "etapa"], [("r1", "AD")])

    def test_arquivo_real_temporario(self):
        with tempfile.TemporaryDirectory() as pasta:
            caminho = Path(pasta) / "relatorio.xlsx"
            wb = Workbook()
            ws = wb.active
            ws.append(["RETRANCA", "STÁTUS"])
            ws.append(["r1", "Descritas"])
            ws.append(["r2", "Conferidas"])
            wb.save(caminho)
            resultado = auditar_arquivo_excel(caminho)
            self.assertEqual(resultado["com_retranca"], 2)
            self.assertEqual(resultado["status"]["CONFERIDAS"], 1)


if __name__ == "__main__":
    unittest.main()
