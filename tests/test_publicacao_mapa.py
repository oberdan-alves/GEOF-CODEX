import json
import tempfile
import unittest
from contextlib import ExitStack
from pathlib import Path
from unittest.mock import patch

import fitz
from openpyxl import load_workbook
import AdjudicacaoOrdem as adj
from test_adjudicacao_pesquisas import item
from test_tela_pesquisa_item import Widget, Var


class PublicacaoMapaTest(unittest.TestCase):
    def test_cinco_publicacoes_rodape_editado_e_criterio_no_pdf(self):
        casos = [((1, 10, 10), 'MÉDIA'), ((10, 20, 100), 'MEDIANA'), ((10, 10, 10), 'MÉDIA = MEDIANA')]
        with tempfile.TemporaryDirectory() as pasta:
            for precos, esperado in casos:
                it = item(precos=precos, pesquisas=(None, None))
                adj.recalcular_referencias(it)
                self.assertEqual(adj.criterio_referencia(it), esperado)
                for publicacao in adj.OPCOES_PUBLICACAO:
                    with self.subTest(publicacao=publicacao, criterio=esperado):
                        arquivo = Path(pasta) / 'mapa.pdf'
                        adj.gerar_pdf_mapa_comparativo(str(arquivo), '123', '2026', [it],
                            publicacao_rotulo=publicacao, sei_publicacao='SEI-ANTIGO',
                            texto_publicacao='Texto editado para o Diário Oficial — SEI NOVO')
                        with fitz.open(arquivo) as pdf:
                            texto = ' '.join(pdf[0].get_text().split())
                            self.assertIn(f'123/2026 – {publicacao}', texto)
                            self.assertIn(f'Preço de referência — {esperado}', texto)
                            self.assertIn('Texto editado para o Diário Oficial — SEI NOVO', texto)
                            self.assertNotIn('SEI-ANTIGO', texto)
                            spans = [s for b in pdf[0].get_text('dict')['blocks'] if 'lines' in b
                                     for l in b['lines'] for s in l['spans']]
                            self.assertTrue(any('Texto editado' in s['text'] and s['color'] == 0xff0000 for s in spans))
                excel = Path(pasta) / 'mapa.xlsx'
                adj.gerar_excel_mapa_comparativo(excel, '123', '2026', [it], {'publicacao_rotulo': '3ª PUBLICAÇÃO'})
                wb = load_workbook(excel)
                self.assertIn('3ª PUBLICAÇÃO', wb['Adjudicação']['A5'].value)
                self.assertIn('MÉDIA', wb['Adjudicação']['K12'].value)
                self.assertIn('MEDIANA', wb['Adjudicação']['K12'].value)
                wb.close()

    def test_tela_salva_e_reabre_publicacao_sem_alterar_outros_dados(self):
        class App(adj.AdjudicacaoOrdemMixin):
            def _caminho_arquivo_mapa_empresas(self, *args):
                return self.arquivo
            def wait_window(self, janela):
                self.exercitar()

        with tempfile.TemporaryDirectory() as pasta, ExitStack() as stack:
            for nome in ('CTkToplevel', 'CTkScrollableFrame', 'CTkLabel', 'CTkTextbox', 'CTkButton', 'CTkOptionMenu'):
                stack.enter_context(patch.object(adj.ctk, nome, Widget))
            stack.enter_context(patch.object(adj.tk, 'StringVar', Var))
            app = App()
            app._pasta_cotacao_atual, app._codigo_cotacao_atual = pasta, '123'
            app.arquivo = str(Path(pasta) / 'mapa.json')
            Path(app.arquivo).write_text(json.dumps({'fornecedores': {'A': {}}, 'decisoes_itens': {'1': {'cancelado': True}}}))

            def controles():
                seletor = next(w for w in Widget.criados if 'values' in w.kwargs)
                rodape = next(w for w in Widget.criados if w.kwargs.get('height') == 80)
                botao = next(w for w in Widget.criados if w.kwargs.get('text') == 'Gerar Mapa Comparativo')
                return seletor, rodape, botao

            def editar():
                seletor, rodape, botao = controles()
                self.assertEqual(seletor.kwargs['values'], adj.OPCOES_PUBLICACAO)
                self.assertEqual(rodape.get(), adj.TEXTO_PUBLICACAO_PADRAO)
                seletor.kwargs['variable'].set('4ª PUBLICAÇÃO')
                rodape.delete('1.0', 'end')
                rodape.insert('1.0', 'Publicação personalizada SEI 999')
                botao.kwargs['command']()

            Widget.criados = []
            app.exercitar = editar
            resultado = app._perguntar_dados_finalizacao()
            self.assertEqual(resultado['publicacao_rotulo'], '4ª PUBLICAÇÃO')
            self.assertEqual(resultado['texto_publicacao'], 'Publicação personalizada SEI 999')
            salvo = json.loads(Path(app.arquivo).read_text())
            self.assertEqual(salvo['fornecedores'], {'A': {}})
            self.assertTrue(salvo['decisoes_itens']['1']['cancelado'])

            def reabrir():
                seletor, rodape, botao = controles()
                self.assertEqual(seletor.kwargs['variable'].get(), '4ª PUBLICAÇÃO')
                self.assertEqual(rodape.get(), 'Publicação personalizada SEI 999')
                botao.kwargs['command']()

            Widget.criados = []
            app.exercitar = reabrir
            app._perguntar_dados_finalizacao()


if __name__ == '__main__':
    unittest.main()
