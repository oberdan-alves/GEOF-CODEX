import copy
from contextlib import ExitStack
from decimal import Decimal
import random
import unittest
from unittest.mock import patch

import AdjudicacaoOrdem as adj
from test_tela_pesquisa_item import Widget, Var


class SimulacaoReferenciaTest(unittest.TestCase):
    def test_exemplo_e_preco_limite(self):
        dados = adj.simular_referencia([11682.85, 2950, 9213.50], 0)
        self.assertAlmostEqual(float(dados['referencia']), 7948.783333, places=5)
        self.assertEqual(dados['limite'], Decimal('6081.75'))
        self.assertFalse(dados['aceita'])
        dados = adj.simular_referencia([7900, 2950, 9213.50], 0)
        self.assertAlmostEqual(float(dados['referencia']), 6687.833333, places=5)
        self.assertFalse(dados['aceita'])
        self.assertTrue(adj.simular_referencia([6081.75, 2950, 9213.50], 0)['aceita'])
        self.assertFalse(adj.simular_referencia([6081.76, 2950, 9213.50], 0)['aceita'])

    def test_teto_respeita_media_e_mediana_com_pares_e_impares(self):
        rng = random.Random(52)
        for quantidade in range(1, 10):
            for _ in range(30):
                fixos = [Decimal(rng.randint(1, 10000))/100 for _ in range(quantidade)]
                limite = adj.simular_referencia([1000] + fixos, 0)['limite']
                self.assertTrue(adj.simular_referencia([limite] + fixos, 0)['aceita'])
                self.assertFalse(adj.simular_referencia([limite + Decimal('.01')] + fixos, 0)['aceita'])
        # A média dos demais é alta, mas a mediana impõe um teto menor.
        self.assertEqual(adj.simular_referencia([2000, 10, 20, 1000], 0)['limite'], Decimal('20.00'))

    def test_exclusoes_vazios_e_valores_invalidos(self):
        self.assertEqual(adj.simular_referencia([100, None, 30], 0)['limite'], Decimal('30.00'))
        self.assertIsNone(adj.simular_referencia([100], 0)['limite'])
        self.assertIsNone(adj.simular_referencia([None, 100], 0)['aceita'])
        self.assertEqual(adj.simular_referencia([None], 0)['quantidade'], 0)
        for valor in (-1, 0, float('nan'), float('inf')):
            with self.assertRaises(ValueError):
                adj.simular_referencia([valor, 10], 0)

    def test_tela_edita_adiciona_retira_e_fecha_sem_alterar_item(self):
        class GradeWidget(Widget):
            def grid(self, **kwargs):
                self.grid_config = kwargs
            def grid_columnconfigure(self, *args, **kwargs):
                pass

        item = dict(numero_item=3, codigo='VME35', vencedor_idx=0,
                    propostas=[dict(fornecedor='ASTEM', unitario=11682.85)],
                    pesquisas=[2950, 9213.50])
        original = copy.deepcopy(item)
        Widget.criados = []
        parent = GradeWidget()
        with ExitStack() as stack:
            for nome in ('CTkToplevel', 'CTkFrame', 'CTkScrollableFrame', 'CTkLabel',
                         'CTkEntry', 'CTkButton', 'CTkCheckBox', 'CTkRadioButton'):
                stack.enter_context(patch.object(adj.ctk, nome, GradeWidget))
            stack.enter_context(patch.object(adj.tk, 'IntVar', Var))
            stack.enter_context(patch.object(adj.tk, 'BooleanVar', Var))
            stack.enter_context(patch.object(adj, 'maximizar_janela'))
            app = adj.AdjudicacaoOrdemMixin()
            app._abrir_simulador_referencia(parent, item, item['pesquisas'], ['NF', 'Painel'])

            def clicar(texto):
                next(w for w in Widget.criados if w.kwargs.get('text') == texto).kwargs['command']()

            campos = [w for w in Widget.criados if w.kwargs.get('width') == 170]
            clicar('TESTAR PREÇO LIMITE')
            self.assertEqual(campos[0].get(), '6.081,75')
            campos[1].delete(0, 'end')
            campos[1].insert(0, '4.000,00')
            clicar('CALCULAR SIMULAÇÃO')
            clicar('ADICIONAR VALOR')
            campos = [w for w in Widget.criados if w.kwargs.get('width') == 170]
            campos[-1].insert(0, '5.000,00')
            checks = [w for w in Widget.criados if 'variable' in w.kwargs and 'value' not in w.kwargs]
            checks[2].kwargs['variable'].set(False)
            clicar('CALCULAR SIMULAÇÃO')
            clicar('FECHAR SIMULAÇÃO')
        self.assertEqual(item, original)


if __name__ == '__main__':
    unittest.main()
