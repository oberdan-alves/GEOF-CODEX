import tempfile
import unittest
from pathlib import Path

import fitz
from Cotacao import gerar_pdf_pedido_cotacao, TEXTO_OBSERVACAO_NECFM
from GeradorCotacao import CONTATOS_ORDEM_PADRAO


class ModelosPedidoTest(unittest.TestCase):
    def gerar(self, pasta, tipo, origem, contatos=None, itens=None):
        caminho = Path(pasta) / f'{tipo}-{origem.replace("/", "-")}.pdf'
        gerar_pdf_pedido_cotacao(str(caminho), '678', '2026', 'CLAYTON',
            '00060-00522943/2026-12', origem, tipo, itens or [dict(ITEM=1.0, CODIGO='5744',
            DESCRICAO='CARVEDILOL COMPRIMIDO 12,5 MG\nRegistro ANVISA e código GGREM.', QTD=120000, UN='CP')],
            contatos or CONTATOS_ORDEM_PADRAO, 'LOYANI KATRINA CABRAL IPAC', '1.661.866-1')
        return fitz.open(caminho)

    def test_cinco_modelos_conteudo_diagramacao_e_fontes(self):
        modelos = [('APS', 'NFH', 'NLF/GAOAPS'), ('APS', 'NECFM', 'NECFM/GAOESP'),
                   ('APS INV', 'NPDOC', 'NPDOC/APS'), ('HRG', 'NUAL', 'NUAL/HRG'),
                   ('HRG INV', 'NPDOC', 'NPDOC/HRG')]
        with tempfile.TemporaryDirectory() as pasta:
            for tipo, origem, nucleo in modelos:
                with self.subTest(tipo=tipo, origem=origem), self.gerar(pasta, tipo, origem) as pdf:
                    self.assertEqual(len(pdf), 1)
                    pagina = pdf[0]
                    texto = ' '.join(pagina.get_text().split())
                    for trecho in ('678/2026', '00060-00522943/2026-12', '10 -', '10.1 - CNPJ:', '11 -', nucleo, '12.1 -', '12.8 -', '13.5 -'):
                        self.assertIn(trecho, texto)
                    self.assertTrue(pagina.get_images())
                    self.assertLess(pagina.search_for('Publ. por:')[0].x0, pagina.search_for('SEI:')[0].x0)
                    self.assertAlmostEqual(pagina.search_for('Publ. por:')[0].y0, pagina.search_for('SEI:')[0].y0, delta=1)
                    self.assertAlmostEqual(pagina.search_for('12.3 -')[0].y0, pagina.search_for('12.7 -')[0].y0, delta=1)
                    self.assertLess(pagina.search_for('12.3 -')[0].x0, pagina.search_for('12.7 -')[0].x0)
                    spans = [s for b in pagina.get_text('dict')['blocks'] if 'lines' in b for l in b['lines'] for s in l['spans']]
                    self.assertTrue(any(s['font'] == 'Times-Roman' for s in spans))
                    self.assertTrue(any(s['font'] == 'Times-Bold' for s in spans))
                    for span in spans:
                        self.assertGreaterEqual(span['bbox'][0], 0)
                        self.assertLessEqual(span['bbox'][2], pagina.rect.width + 1)
                    if 'HRG' in tipo:
                        self.assertIn('CNPJ: 00.394.700/0006-12', texto)
                        self.assertIn('PDPAS/HRG - GEOF', texto)
                    else:
                        self.assertIn('CNPJ: 00.394.700/0051-77', texto)
                        self.assertIn('PDPAS / APS-SUL - GEOF', texto)
                        self.assertNotIn('PDPAS/HRG - GEOF', texto)
                    if 'INV' in tipo:
                        self.assertIn('12.7 - O preço do medicamento deve estar abaixo do Preço Fábrica da CMED.', texto)
                    else:
                        self.assertIn('12.7 - Medicamento informar o GGREM e Nº Anvisa.', texto)
                        if origem != 'NECFM':
                            self.assertIn('14- O contratado deverá manter', texto)
                    if origem == 'NECFM':
                        for linha in TEXTO_OBSERVACAO_NECFM:
                            self.assertIn(linha, texto)
                        self.assertTrue(any(s['color'] == 0xff0000 for s in spans))
                        self.assertNotIn('Chefe do Núcleo de Engenharia Clínica e Física Médica/HRG', texto)
                    if tipo == 'APS INV':
                        self.assertIn('PESQUISA DE PREÇOS PARA AQUISIÇÃO POR CONTRAPARTIDA', texto)

    def test_farmacia_aps_aliases_e_contato_personalizado(self):
        with tempfile.TemporaryDirectory() as pasta:
            for origem in ('NFL', 'NFH', 'NLF', 'NFL/GAOAPS-SUL'):
                with self.gerar(pasta, 'APS', origem, contatos=dict(CONTATOS_ORDEM_PADRAO, NFL='CONTATO PERSONALIZADO')) as pdf:
                    texto = ' '.join(pdf[0].get_text().split())
                    self.assertIn('NLF/GAOAPS', texto)
                    self.assertIn('CONTATO PERSONALIZADO', texto)

    def test_varios_itens_preservam_paginacao_e_observacao_ajustada(self):
        itens = [dict(ITEM=i, CODIGO=f'C{i}', DESCRICAO='Descrição técnica detalhada. ' * 15,
                      QTD=1, UN='UN') for i in range(1, 16)]
        itens[0]['_observacao_destaque'] = 'OBSERVAÇÃO PERSONALIZADA DO PEDIDO'
        with tempfile.TemporaryDirectory() as pasta, self.gerar(pasta, 'HRG', 'NUAL', itens=itens) as pdf:
            self.assertGreater(len(pdf), 1)
            texto = ' '.join(' '.join(p.get_text() for p in pdf).split())
            self.assertIn('OBSERVAÇÃO PERSONALIZADA DO PEDIDO', texto)
            self.assertIn('C15', texto)
            self.assertIn('14- O contratado deverá manter', texto)


if __name__ == '__main__':
    unittest.main()
