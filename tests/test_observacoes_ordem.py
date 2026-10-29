import json
import tempfile
import unittest
from contextlib import ExitStack
from pathlib import Path
from unittest.mock import Mock, patch

import fitz
import AdjudicacaoOrdem as adj
from test_tela_pesquisa_item import Widget


class ObservacoesOrdemTest(unittest.TestCase):
    def test_texto_destacado_entre_total_e_dados_nf_nas_quatro_variantes(self):
        with tempfile.TemporaryDirectory() as pasta:
            for tipo in ('APS', 'APS INV', 'HRG', 'HRG INV'):
                with self.subTest(tipo=tipo):
                    arquivo = Path(pasta) / 'ordem.pdf'
                    adj.gerar_pdf_ordem_compra(str(arquivo), '1', '2026', 'TR', 'NFH',
                        {'nome': 'EMPRESA'}, [], tipo, '', '', '', observacoes='Entrega <agendada> & confirmada.\nAvisar antes de entregar.')
                    with fitz.open(arquivo) as pdf:
                        texto = ' '.join(' '.join(p.get_text() for p in pdf).split())
                        self.assertLess(texto.index('TOTAL'), texto.index('OBSERVAÇÕES:'))
                        self.assertLess(texto.index('Avisar antes de entregar.'), texto.index('Dados para emissão de Nota Fiscal'))
                        self.assertIn('Entrega <agendada> & confirmada.', texto)
                        spans = [s for p in pdf for b in p.get_text('dict')['blocks'] if 'lines' in b
                                 for l in b['lines'] for s in l['spans']]
                        self.assertTrue(any('Entrega' in s['text'] and s['font'] == 'Helvetica-Bold'
                                            and s['color'] == 0x8b0000 for s in spans))
                    adj.gerar_pdf_ordem_compra(str(arquivo), '1', '2026', 'TR', 'NFH',
                                              {'nome': 'EMPRESA'}, [], tipo, '', '', '', observacoes='')
                    with fitz.open(arquivo) as pdf:
                        self.assertNotIn('OBSERVAÇÕES:', ' '.join(p.get_text() for p in pdf))

    def test_dialogo_salva_reabre_e_cancelamento_preserva_texto(self):
        class App(adj.AdjudicacaoOrdemMixin):
            def _caminho_arquivo_mapa_empresas(self, *args):
                return self.arquivo
            def wait_window(self, janela):
                self.exercitar(janela)

        with tempfile.TemporaryDirectory() as pasta, ExitStack() as stack:
            for nome in ('CTkToplevel', 'CTkFrame', 'CTkLabel', 'CTkTextbox', 'CTkButton'):
                stack.enter_context(patch.object(adj.ctk, nome, Widget))
            stack.enter_context(patch.object(adj, 'centralizar_janela'))
            arquivo = Path(pasta) / 'mapa.json'
            arquivo.write_text(json.dumps({'fornecedores': {'A': {}}, 'dados_mapa': {'objeto': 'Preservar'}}))
            app = App()
            app.arquivo = str(arquivo)
            app._pasta_cotacao_atual, app._codigo_cotacao_atual = pasta, '1'
            texto = 'Observação exclusiva desta cotação.\nSegunda linha.'

            def editar(janela):
                campo = next(w for w in Widget.criados if w.kwargs.get('wrap') == 'word')
                campo.insert('1.0', texto)
                next(w for w in Widget.criados if w.kwargs.get('text') == 'GERAR ORDENS').kwargs['command']()

            Widget.criados = []
            app.exercitar = editar
            self.assertEqual(app._perguntar_observacoes_ordem(), texto)
            salvo = json.loads(arquivo.read_text())
            self.assertEqual(salvo['observacoes_ordem'], texto)
            self.assertEqual(salvo['fornecedores'], {'A': {}})
            self.assertEqual(salvo['dados_mapa'], {'objeto': 'Preservar'})

            def cancelar(janela):
                campo = next(w for w in Widget.criados if w.kwargs.get('wrap') == 'word')
                self.assertEqual(campo.get(), texto)
                campo.insert('1.0', 'Edição cancelada')
                next(w for w in Widget.criados if w.kwargs.get('text') == 'CANCELAR').kwargs['command']()

            Widget.criados = []
            app.exercitar = cancelar
            self.assertIsNone(app._perguntar_observacoes_ordem())
            self.assertEqual(json.loads(arquivo.read_text()), salvo)

    def test_cancelar_ordem_nao_gera_documentos(self):
        app = adj.AdjudicacaoOrdemMixin()
        app._montar_dados_ordens_ativas = Mock(return_value={})
        app._perguntar_observacoes_ordem = Mock(return_value=None)
        with patch.object(adj, 'gerar_pdf_ordem_compra') as gerar:
            app.acao_ordem()
            gerar.assert_not_called()


if __name__ == '__main__':
    unittest.main()
