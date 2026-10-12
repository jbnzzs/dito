"""Testes do alerta editorial G27, sem acesso ao banco."""
import unittest
from types import SimpleNamespace
from core.services.origem_service import origem_g27


class OrigemG27Tests(unittest.TestCase):
    def imagem(self, retranca='p013_cie_tela', projeto='Teste G28', editorial=None):
        return SimpleNamespace(
            retranca=retranca,
            projeto=SimpleNamespace(nome=projeto),
            dados_fotoweb_originais=(
                {'__projeto_editorial_pdf': editorial} if editorial else {}
            ),
        )

    def test_projeto_g27(self):
        self.assertTrue(origem_g27(self.imagem(projeto='G27')))
        self.assertTrue(origem_g27(self.imagem(projeto='Teste G27')))

    def test_origem_no_snapshot(self):
        self.assertTrue(origem_g27(self.imagem(editorial='G27_E002_EF1')))

    def test_retranca(self):
        self.assertTrue(origem_g27(self.imagem(retranca='g27_e002_4p012m_mp')))

    def test_g28_e_livro_ef1_nao_comprovam_g27(self):
        self.assertFalse(origem_g27(self.imagem(retranca='g28_e002_4p013m_mp')))
        self.assertFalse(origem_g27(self.imagem(retranca='g270_e002_6p013m_mp')))


if __name__ == '__main__':
    unittest.main()
