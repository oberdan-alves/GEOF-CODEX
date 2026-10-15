import tempfile
import unittest
from pathlib import Path
import fitz
from Mapa import montar_ranking_itens, gerar_pdf_ranking


class MapaPrevioTest(unittest.TestCase):
    def test_agrupamento_descricao_vermelha_e_colocacoes(self):
        fornecedores = {nome: {'itens': [dict(numero_item=7, codigo='CAN11',
                descricao='CONSERTO APARELHO DE ANESTESIA COM MÓVEL', quantidade=1.0,
                unitario=preco, total=preco, prazo=20)]}
                for nome, preco in [('WORLD', 5000), ('BIOLOG (396)', 2854)]}
        ranking = montar_ranking_itens(fornecedores, {})
        self.assertEqual(ranking[0][1]['numero_item'], 7)
        with tempfile.TemporaryDirectory() as pasta:
            arquivo = Path(pasta)/'mapa.pdf'
            gerar_pdf_ranking(str(arquivo), '109', '2026', ranking)
            with fitz.open(arquivo) as doc:
                texto = ' '.join(doc[0].get_text().split())
                spans = [s for b in doc[0].get_text('dict')['blocks'] if 'lines' in b
                         for l in b['lines'] for s in l['spans']]
            self.assertIn('MAPA COMPARATIVO PRÉVIO RESUMIDO', texto)
            self.assertIn('7 CAN11 CONSERTO APARELHO DE ANESTESIA COM MÓVEL', texto)
            self.assertIn('1º COLOCADO BIOLOG (396) 1', texto)
            self.assertIn('2º COLOCADO WORLD 1', texto)
            self.assertEqual(texto.count('CAN11'), 1)
            self.assertTrue(any('CONSERTO APARELHO' in s['text'] and s['color'] == 0xff0000 for s in spans))

    def test_descricao_longa_nao_ultrapassa_margem(self):
        ranking = [('ABC', dict(numero_item=1, descricao='DESCRIÇÃO LONGA ' * 200, propostas=[]))]
        with tempfile.TemporaryDirectory() as pasta:
            arquivo = Path(pasta)/'mapa.pdf'
            gerar_pdf_ranking(str(arquivo), '109', '2026', ranking)
            with fitz.open(arquivo) as doc:
                self.assertIn('...', doc[0].get_text())
                for b in doc[0].get_text('blocks'):
                    self.assertLessEqual(b[2], doc[0].rect.width)
