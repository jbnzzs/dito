
import unittest

from core.services.importacao_service import (
    classificar_status_fotoweb,
)


class TestMapeamentoFotoWeb(unittest.TestCase):

    def test_status_importaveis(self):
        casos = {
            "1.DESCREVER": "Liberado para descrição",
            "2.DESCREVENDO": "Descrevendo",
            "3.DESCRITA": "Descrito",
            "4.CONFERINDO": "Conferindo",
            "5.CONFERIDA": "Conferido",
            "6.REVISANDO": "Avaliando",
            "7.REVISADA": "Avaliando",
        }

        for origem, destino in casos.items():
            with self.subTest(origem=origem):
                acao, resultado, _ = (
                    classificar_status_fotoweb(origem)
                )
                self.assertEqual(acao, "importar")
                self.assertEqual(resultado, destino)

    def test_variacoes(self):
        casos = {
            "descritas": "Descrito",
            "CONFERIDAS": "Conferido",
            "revisadas": "Avaliando",
            "Em descrição": "Descrevendo",
            "Descrição concluída": "Descrito",
        }

        for origem, destino in casos.items():
            with self.subTest(origem=origem):
                acao, resultado, _ = (
                    classificar_status_fotoweb(origem)
                )
                self.assertEqual((acao, resultado),
                                 ("importar", destino))

    def test_status_ignorados(self):
        for origem in (
            "0.EXCLUIDA",
            "8.COTEJANDO",
            "9.COTEJADA",
        ):
            with self.subTest(origem=origem):
                acao, destino, _ = (
                    classificar_status_fotoweb(origem)
                )
                self.assertEqual(acao, "ignorar")
                self.assertIsNone(destino)

    def test_status_desconhecido(self):
        acao, destino, _ = (
            classificar_status_fotoweb("STATUS NOVO")
        )
        self.assertEqual(acao, "revisar")
        self.assertIsNone(destino)

    def test_codigo_incorreto(self):
        acao, destino, _ = (
            classificar_status_fotoweb("7.DESCRITA")
        )
        self.assertEqual(acao, "revisar")
        self.assertIsNone(destino)

    def test_status_vazio(self):
        acao, destino, _ = classificar_status_fotoweb(None)
        self.assertEqual(acao, "revisar")
        self.assertIsNone(destino)


if __name__ == "__main__":
    unittest.main()
