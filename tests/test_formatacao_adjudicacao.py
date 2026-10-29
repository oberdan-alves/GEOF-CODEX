import copy
import tempfile
import unittest
from pathlib import Path

import fitz
import AdjudicacaoOrdem as adj
from test_adjudicacao_pesquisas import item


class FormatacaoAdjudicacaoTest(unittest.TestCase):
    def gerar(self, pasta, tipo, itens):
        caminho = Path(pasta) / f'{tipo}.pdf'
        adj.gerar_pdf_mapa_comparativo(str(caminho), '987', '2027', itens,
            objeto_titulo='OBJETO DINÂMICO DA COTAÇÃO', texto_pesquisa='JUSTIFICATIVA PERSONALIZADA\nSegunda linha.',
            observacoes='OBSERVAÇÃO PERSONALIZADA', diretor_nome='DIRETOR CONFIGURADO',
            diretor_matricula='123', sei_publicacao='SEI-DINÂMICO', tipo_planilha=tipo,
            titulos_pesquisa=['Fonte externa atual 1', 'Fonte externa atual 2'])
        return fitz.open(caminho)

    def test_quatro_variantes_preservam_dados_e_usam_hierarquia_dos_modelos(self):
        it = item()
        it['propostas'][0]['quantidade'] = 1.25
        adj.recalcular_referencias(it)
        original = copy.deepcopy(it)
        with tempfile.TemporaryDirectory() as pasta:
            for tipo in ('APS', 'APS INV', 'HRG', 'HRG INV'):
                with self.subTest(tipo=tipo), self.gerar(pasta, tipo, [it]) as pdf:
                    self.assertEqual(len(pdf), 1)
                    texto = ' '.join(pdf[0].get_text().split())
                    for esperado in ('987/2027', '1,25', 'OBJETO DINÂMICO', 'JUSTIFICATIVA PERSONALIZADA',
                                     'OBSERVAÇÃO PERSONALIZADA', 'DIRETOR CONFIGURADO', 'SEI-DINÂMICO',
                                     'Fonte externa atual 1', 'Fonte externa atual 2', 'VME01'):
                        self.assertIn(esperado, texto)
                    spans = [s for b in pdf[0].get_text('dict')['blocks'] if 'lines' in b
                             for linha in b['lines'] for s in linha['spans']]
                    self.assertTrue(any(s['font'] == 'Times-Roman' for s in spans))
                    self.assertTrue(any(s['font'] == 'Times-Bold' for s in spans))
                    self.assertEqual(len(pdf[0].search_for('R$ 5.650,00')), 2)
                    self.assertIn('PDPAS/HRG' if 'HRG' in tipo else 'PDPAS / APS-SUL', texto)
        self.assertEqual(it, original)

    def test_destaque_cobre_nome_sem_quantidade_valores_prazo_ou_marca(self):
        for pos in (0, 1, 2):
            it = item()
            it['vencedor_idx'] = pos
            for i, p in enumerate(it['propostas']):
                p['fornecedor'] = f'EMPRESA {i}'
            adj.recalcular_referencias(it)
            with self.subTest(pos=pos), tempfile.TemporaryDirectory() as pasta, self.gerar(pasta, 'HRG', [it]) as pdf:
                pagina = pdf[0]
                nome = pagina.search_for(f'EMPRESA {pos}')[0]
                fundos = [fitz.Rect(d['rect']) for d in pagina.get_drawings()
                          if d.get('fill') and all(abs(c - 217/255) < .01 for c in d['fill'])]
                destaque = next(r for r in fundos if r.contains((nome.tl + nome.br) / 2))
                self.assertLess(destaque.width, pagina.rect.width * .20)
                self.assertLess(destaque.height, 20)
                outros = [pagina.search_for(f'EMPRESA {i}')[0] for i in range(3) if i != pos]
                for outro in outros:
                    self.assertFalse(destaque.intersects(outro))
                self.assertLess(destaque.y1, pagina.search_for('Prazo (dias)')[0].y0)

    def test_muitos_itens_paginam_sem_reduzir_fonte_ou_separar_bloco(self):
        itens = [item(numero=i) for i in range(1, 26)]
        for it in itens:
            adj.recalcular_referencias(it)
        with tempfile.TemporaryDirectory() as pasta, self.gerar(pasta, 'APS', itens) as pdf:
            self.assertGreater(len(pdf), 1)
            texto = ' '.join(' '.join(p.get_text() for p in pdf).split())
            for i in range(1, 26):
                self.assertIn(f'VME{i:02}', texto)
            self.assertIn('ADJUDICAÇÃO E HOMOLOGAÇÃO', texto)
            self.assertIn('JUSTIFICATIVA PERSONALIZADA', texto)
            for pagina in pdf:
                codigos = [w for w in pagina.get_text('words') if w[4].startswith('VME')]
                if codigos:
                    self.assertIn('Item /', pagina.get_text())
                    self.assertGreater(codigos[0][1], pagina.search_for('Código')[0].y1)


if __name__ == '__main__':
    unittest.main()
