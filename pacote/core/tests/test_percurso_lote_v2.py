"""Testes independentes do banco: python -m unittest discover -s core/tests -p '*percurso_lote_v2.py'"""
import unittest
from importlib.util import module_from_spec, spec_from_file_location
from pathlib import Path

CAMINHO = Path(__file__).resolve().parents[1] / 'services' / 'percurso_lote_service.py'
spec = spec_from_file_location('dito_percurso_persistente_v2', CAMINHO)
servico = module_from_spec(spec)
spec.loader.exec_module(servico)


class PercursoV2Testes(unittest.TestCase):
    def test_ordem_fixa_acrescenta_sem_reordenar(self):
        self.assertEqual(servico._sincronizar_ordem([3, 1, 2], [4, 2, 5]), [3, 1, 2, 4, 5])

    def test_proximo_respeita_ordenacao_inicial(self):
        self.assertEqual(servico._selecionar_proxima([8, 3, 5], 8, [3, 5]), 3)

    def test_proximo_ignora_prontas_ao_procurar_tarefa_nova(self):
        self.assertEqual(servico._selecionar_proxima([8, 3, 5], 8, [5]), 5)

    def test_voltar_e_avancar_revisitam_mesmo_se_pronta(self):
        self.assertEqual(
            servico._selecionar_proxima([8, 3, 5], 3, [5], [8, 3, 5], 1, [8, 3, 5]),
            5,
        )

    def test_retorno_nao_pula_imagem_inacabada(self):
        self.assertEqual(servico._selecionar_proxima([8, 3, 5], 3, [3, 5]), 5)

    def test_ciclo_nao_retorna_a_mesma_imagem(self):
        self.assertIsNone(servico._selecionar_proxima([8], 8, [8]))

    def test_encontrar_pendente_ao_retomar_sem_historico(self):
        self.assertEqual(servico._selecionar_proxima([1, 2, 3], 1, [3]), 3)


if __name__ == '__main__':
    unittest.main()
