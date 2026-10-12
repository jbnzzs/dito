"""Testes puros, sem banco, para a lógica de links de PDF do Dito!."""
import unittest

from core.services.excel_service import _valor_excel
from core.services import sharepoint_service as sp


class SharepointServiceTests(unittest.TestCase):
    def test_primeiro_valor_preenchido(self):
        self.assertEqual(_valor_excel({"a": "", "b": "9"}, "a", "b"), "9")

    def test_normalizacao_componente(self):
        self.assertEqual(sp._codigo_componente_pdf_sharepoint("Educação física"), "EFIS")
        self.assertEqual(sp._codigo_componente_pdf_sharepoint("Matemática"), "MAT")

    def test_pdf_matematica(self):
        url = sp._montar_url_pdf_metadados_dito("G28_E002_EF2", "p013", "MAT", "9")
        self.assertEqual(
            url,
            "https://ensinolivre.sharepoint.com/sites/G25Digital/"
            "Shared%20Documents/g28/PDFs/MAT/g28_e002_9p013m_mp.pdf?web=1",
        )

    def test_ed_fisica_ano_6_conhecido(self):
        self.assertTrue(sp._montar_url_pdf_metadados_dito("G28_E002_EF2", "p013", "EFIS", "6"))

    def test_ed_fisica_ano_9_ainda_nao_mapeado(self):
        self.assertEqual(sp._montar_url_pdf_metadados_dito("G28_E002_EF2", "p013", "EFIS", "9"), "")

    def test_conflito_retranca(self):
        url, motivo = sp._url_pdf_dados_fotoweb(
            "G28_E002_EF2", "p013", {"componente": "Matemática", "volume": "9"}, "p013_cie_tecla",
        )
        self.assertEqual((url, motivo), ("", "conflito_excel_retranca"))

    def test_volume_ambiguo(self):
        url, motivo = sp._url_pdf_dados_fotoweb(
            "G28_E002_EF2", "p013", {"componente": "Matemática", "volume": "2029"}
        )
        self.assertEqual((url, motivo), ("", "ano_ausente_ou_ambiguo"))

    def test_retranca_nao_inventa_projeto(self):
        self.assertEqual(sp._montar_url_pdf_sharepoint("p013_tecla_1_fim", "Matemática"), "")


if __name__ == "__main__":
    unittest.main()
