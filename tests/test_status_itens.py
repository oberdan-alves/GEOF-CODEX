import json
import tempfile
import unittest
from contextlib import ExitStack
from pathlib import Path
from unittest.mock import Mock, patch

from openpyxl import load_workbook
import fitz
import AdjudicacaoOrdem as adj
from test_adjudicacao_pesquisas import item
from test_tela_pesquisa_item import Widget, Var


class StatusItensTest(unittest.TestCase):
    def test_textos_justificativa_e_menos_tres_valores_no_relatorio(self):
        it = item(precos=(100,), pesquisas=(None, None))
        texto = adj.montar_texto_relatorio_pesquisa([it], '1', '2026', {
            'texto_pesquisa': adj.TEXTO_PESQUISA_PADRAO + '\nComplemento manual.'})
        aviso = ('Atenção: conjunto com menos de três valores. '
                 'A justificativa e os documentos de suporte fazem parte do processo.')
        self.assertIn(aviso, texto)
        self.assertNotIn('devem ser conferidos pelo responsável', texto)
        topico4 = texto.split('4. PESQUISA DE PREÇOS E JUSTIFICATIVA')[1].split('5. OBSERVAÇÕES')[0]
        paragrafo = ('Em função do que se apresenta na planilha acima, das necessidades expostas no Termo de Referência '
                     'e que o(s) produto(s) ofertado(s) atende(m) ao(s) requisito(s) de melhor proposta, '
                     'prevista no edital referente à presente cotação.')
        self.assertIn(paragrafo, topico4)
        self.assertNotIn('solicitamos de V.Sª', topico4)
        self.assertIn('Complemento manual.', topico4)
        antigo = adj.TEXTO_PESQUISA_PADRAO + '\n' + aviso.replace('fazem parte do processo.', 'devem ser conferidos pelo responsável.')
        corrigido = adj.corrigir_textos_relatorio(antigo)
        self.assertIn(aviso, corrigido)
        self.assertIn(paragrafo, corrigido)
        self.assertNotIn('solicitamos de V.Sª', corrigido)

    def test_numeros_inteiros_e_conclusao_em_relatorio_novo_e_salvo(self):
        itens = [item(numero=1.0), item(numero=3.0)]
        itens[0]['status_item'] = 'FRACASSADA'
        observacoes = adj.TEXTO_OBSERVACAO_PADRAO.format(itens='1.0, 3.0')
        texto = adj.montar_texto_relatorio_pesquisa(itens, '1', '2026', {'observacoes': observacoes})
        self.assertIn('item(ns) 1, 3, uma vez', texto)
        self.assertIn('ITEM 1 —', texto)
        self.assertIn('ITEM 3 —', texto)
        self.assertIn('das propostas e dos pareceres mencionados acompanham o processo.', texto)
        self.assertNotIn('deve acompanhar o processo', texto)
        self.assertIn('ITEM 1 —', adj.observacoes_status_itens(itens))
        antigo = texto.replace('ITEM 1 —', 'ITEM 1.0 —').replace('item(ns) 1, 3,', 'item(ns) 1.0, 3.0,')
        antigo = antigo.replace('mencionados acompanham', 'mencionados deve acompanhar')
        antigo += '\nComplemento manual: 2.0 mg.'
        atualizado = adj.atualizar_status_relatorio(antigo, texto)
        self.assertIn('item(ns) 1, 3, uma vez', atualizado)
        self.assertNotIn('ITEM 1.0', atualizado)
        self.assertIn('Complemento manual: 2.0 mg.', atualizado)
        self.assertIn('mencionados acompanham o processo.', atualizado)

    def test_ordem_ignora_itens_com_status_salvo(self):
        with tempfile.TemporaryDirectory() as pasta:
            app = adj.AdjudicacaoOrdemMixin()
            app._pasta_cotacao_atual = pasta
            app._preparar_contexto_wizard = Mock(return_value=('propostas', '1', '2026'))
            app._obter_caminho_arquivo = Mock(return_value='matrix.xlsx')
            app.entry_sei = Mock(get=Mock(return_value=''))
            app.saldos_bd_atual = {}
            app.linhas_resultado = [dict(ITEM=i, CODIGO=f'C{i}', DESCRICAO='Serviço', UN='un')
                                    for i in range(1, 6)]
            arquivo = Path(pasta) / 'mapa.json'
            app._caminho_arquivo_mapa_empresas = Mock(return_value=str(arquivo))
            decisoes = {str(i): dict(status_item=status, vencedor_idx=None)
                        for i, status in enumerate(adj.STATUS_ITEM, 1)}
            decisoes['5'] = dict(vencedor_idx=0)
            arquivo.write_text(json.dumps(dict(decisoes_itens=decisoes, fornecedores={
                'A': dict(itens=[dict(codigo=f'C{i}', unitario=10, quantidade=1, total=10)
                                 for i in range(1, 6)])})))
            with patch.object(adj, 'carregar_fornecedores_completo', return_value={}):
                dados = app._montar_dados_ordens_ativas()
            itens = next(iter(dados['ordens_por_fornecedor'].values()))[1]
            self.assertEqual([it['numero_item'] for it, _ in itens], [5])

    def test_textos_relatorio_mapa_e_exclusao_dos_vencedores(self):
        for status in adj.STATUS_ITEM:
            with self.subTest(status=status), tempfile.TemporaryDirectory() as pasta:
                it = item()
                it.update(status_item=status, vencedor_idx=None)
                adj.recalcular_referencias(it)
                observacoes = adj.observacoes_status_itens([it])
                texto = adj.montar_texto_relatorio_pesquisa([it], '1', '2026', {'observacoes': observacoes})
                analise = texto.split('3. ANÁLISE POR ITEM')[1].split('4. PESQUISA')[0]
                self.assertIn(adj.texto_status_item(it), analise)
                self.assertIn(adj.STATUS_ITEM[status], analise)
                self.assertNotIn('Empresa vencedora selecionada:', analise)
                self.assertEqual(adj.montar_declaracao_vencedores([it]), [])
                for outro in adj.STATUS_ITEM:
                    if outro != status:
                        self.assertNotIn(f'DISPENSA DE LICITAÇÃO {outro}', texto)
                arquivo = Path(pasta) / 'mapa.xlsx'
                adj.gerar_excel_mapa_comparativo(arquivo, '1', '2026', [it], {})
                wb = load_workbook(arquivo)
                self.assertEqual(wb['Adjudicação']['B9'].value, adj.texto_status_item(it))
                wb.close()
                pdf = Path(pasta) / 'mapa.pdf'
                adj.gerar_pdf_mapa_comparativo(str(pdf), '1', '2026', [it])
                with fitz.open(pdf) as doc:
                    texto_pdf = ' '.join(' '.join(p.get_text() for p in doc).split())
                self.assertIn(adj.texto_status_item(it), texto_pdf)

    def test_relatorio_editado_atualiza_so_desfecho(self):
        it = item()
        base = adj.montar_texto_relatorio_pesquisa([it], '1', '2026', {})
        editado = base.replace('Conserto ventilador pulmonar', 'Descrição revisada manualmente')
        for status in adj.STATUS_ITEM:
            it['status_item'] = status
            automatico = adj.montar_texto_relatorio_pesquisa([it], '1', '2026', {
                'observacoes': adj.observacoes_status_itens([it])})
            editado = adj.atualizar_status_relatorio(editado, automatico)
            self.assertIn('Descrição revisada manualmente', editado)
            self.assertIn(adj.texto_status_item(it), editado)
            self.assertNotIn('Empresa vencedora selecionada:', editado)
            for outro in adj.STATUS_ITEM:
                if outro != status:
                    self.assertNotIn(f'DISPENSA DE LICITAÇÃO {outro}', editado)
        it['status_item'] = ''
        automatico = adj.montar_texto_relatorio_pesquisa([it], '1', '2026', {})
        editado = adj.atualizar_status_relatorio(editado, automatico)
        self.assertNotIn('DISPENSA DE LICITAÇÃO', editado)
        self.assertIn('Empresa vencedora selecionada:', editado)

    def test_botoes_destaque_persistencia_sem_vencedor_e_item_sem_propostas(self):
        with tempfile.TemporaryDirectory() as pasta, ExitStack() as stack:
            for nome in ('CTkToplevel', 'CTkFrame', 'CTkScrollableFrame', 'CTkLabel',
                         'CTkEntry', 'CTkTextbox', 'CTkCheckBox', 'CTkRadioButton', 'CTkButton'):
                stack.enter_context(patch.object(adj.ctk, nome, Widget))
            stack.enter_context(patch.object(adj.tk, 'IntVar', Var))
            stack.enter_context(patch.object(adj.tk, 'BooleanVar', Var))
            arquivo = Path(pasta) / 'mapa.json'
            arquivo.write_text(json.dumps({'decisoes_itens': {}}))
            primeiro, segundo = item(), item(precos=(), numero=2)
            Widget.criados = []

            def widgets(texto):
                return [w for w in Widget.criados if not w.destruido and texto in w.kwargs.get('text', '')]

            def clicar(texto):
                widgets(texto)[0].kwargs['command']()

            class App(adj.AdjudicacaoOrdemMixin):
                def _caminho_arquivo_mapa_empresas(self, *args):
                    return str(arquivo)
                def wait_window(app, janela):
                    for status in adj.STATUS_ITEM:
                        clicar(status)
                        self.assertEqual(primeiro['status_item'], status)
                        self.assertIsNone(primeiro['vencedor_idx'])
                        botao = next(w for w in widgets(status) if w.kwargs.get('text') == status)
                        self.assertEqual(botao.kwargs['text_color'], '#FFFF80')
                        self.assertIn('<Enter>', botao.kwargs)
                        radio = next(w for w in widgets('colocado') if 'value' in w.kwargs)
                        self.assertEqual(radio.kwargs['variable'].get(), -1)
                        self.assertIn(f'DISPENSA DE LICITAÇÃO {status}', radio.kwargs['text'])
                    decisao = json.loads(arquivo.read_text())['decisoes_itens']['1']
                    self.assertEqual(decisao['status_item'], 'ANULADA')
                    self.assertIsNone(decisao['vencedor_idx'])
                    radio = next(w for w in widgets('colocado') if 'value' in w.kwargs)
                    radio.kwargs['variable'].set(1)
                    radio.kwargs['command']()
                    self.assertEqual(primeiro['status_item'], '')
                    self.assertEqual(primeiro['vencedor_idx'], 1)
                    clicar('ANULADA')
                    clicar('RECALCULAR')
                    self.assertIsNone(primeiro['vencedor_idx'])
                    clicar('CONFIRMAR EMPRESA')
                    clicar('DESERTA')
                    self.assertTrue(widgets('DISPENSA DE LICITAÇÃO DESERTA'))
                    self.assertIsNone(segundo['vencedor_idx'])
                    clicar('LIMPAR DADOS')
                    self.assertEqual(segundo['status_item'], '')
                    clicar('DESERTA')
                    clicar('CONFIRMAR EMPRESA')
            app = App()
            app._pasta_cotacao_atual, app._codigo_cotacao_atual = pasta, '1'
            app._perguntar_pesquisas_precos([primeiro, segundo])
            self.assertEqual(json.loads(arquivo.read_text())['decisoes_itens']['2']['status_item'], 'DESERTA')

    def test_dica_exibe_descricao_ao_passar_mouse(self):
        botao = Mock()
        botao.winfo_rootx.return_value = 20
        botao.winfo_rooty.return_value = 200
        adj.adicionar_dica_status(botao, adj.STATUS_ITEM['FRACASSADA'])
        eventos = {args[0]: args[1] for args, _ in botao.bind.call_args_list}
        eventos['<Enter>']()
        with patch.object(adj.ctk, 'CTkToplevel') as janela, patch.object(adj.ctk, 'CTkLabel') as label:
            botao.after.call_args.args[1]()
            self.assertEqual(label.call_args.kwargs['text'], adj.STATUS_ITEM['FRACASSADA'])
            eventos['<Leave>']()
            janela.return_value.destroy.assert_called_once()


if __name__ == '__main__':
    unittest.main()
