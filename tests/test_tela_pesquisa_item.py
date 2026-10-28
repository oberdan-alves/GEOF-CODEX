"""Exercita os callbacks da tela por item sem depender de um display."""
import unittest
from unittest.mock import patch
from contextlib import ExitStack

import AdjudicacaoOrdem as adj
from test_adjudicacao_pesquisas import item


class Var:
    def __init__(self, master=None, value=None):
        self.value = value
    def get(self):
        return self.value
    def set(self, value):
        self.value = value


class Widget:
    criados = []
    def __init__(self, parent=None, **kwargs):
        self.parent, self.kwargs = parent, kwargs
        self.children, self.valor, self.destruido = [], '', False
        if isinstance(parent, Widget):
            parent.children.append(self)
        self.criados.append(self)
    def pack(self, **kwargs):
        self.pack_config = kwargs
    def bind(self, evento, callback):
        self.kwargs[evento] = callback
    def configure(self, **kwargs):
        self.kwargs.update(kwargs)
    def winfo_children(self):
        return list(self.children)
    def destroy(self):
        self.destruido = True
        for child in list(self.children):
            child.destroy()
        if isinstance(self.parent, Widget):
            self.parent.children.remove(self)
    def insert(self, index, valor):
        self.valor = str(valor)
    def delete(self, *args):
        self.valor = ''
    def get(self, *args):
        return self.valor
    def protocol(self, evento, callback):
        self.kwargs[evento] = callback
    def geometry(self, tamanho):
        self.tamanho = tamanho
    def minsize(self, *args):
        pass
    def title(self, title):
        pass
    def grab_set(self):
        pass
    def after_idle(self, callback):
        pass
    def after(self, delay, callback):
        self.maximizar_agendado = delay


class TelaPesquisaItemTest(unittest.TestCase):
    def setUp(self):
        salvamento = patch.object(adj.AdjudicacaoOrdemMixin, '_salvar_campos_telas', return_value=True)
        self.salvar = salvamento.start()
        self.addCleanup(salvamento.stop)

    def test_recalculo_historico_observacoes_e_todos_colocados(self):
        primeiro = item(pesquisas=(None, None))
        primeiro['descricao'] = 'Descrição completa do serviço ' * 50
        segundo = item(precos=(10, 11, 12, 13, 14, 15, 16), pesquisas=(None, None), numero=2)
        cancelado = item(numero=3)
        cancelado['cancelado'] = True
        Widget.criados = []
        def widgets(texto):
            return [w for w in Widget.criados if not w.destruido and texto in w.kwargs.get('text', '')]
        def clicar(texto):
            widgets(texto)[0].kwargs['command']()
        def campos():
            # Os dois campos com fonte 20 são os valores unitários do item atual.
            return [w for w in Widget.criados if not w.destruido and w.kwargs.get('font') == ('Segoe UI', 20)]
        class App(adj.AdjudicacaoOrdemMixin):
            def wait_window(self, janela):
                self.janela = janela
                self.assertions()
        app = App()
        def exercitar():
            self.assertTrue(widgets('ITEM 1'))
            resumo = next(w for w in Widget.criados if w.kwargs.get('wraplength') == 0)
            self.assertEqual(len(resumo.kwargs['text'].splitlines()), 1)
            self.assertIn('ITEM 1 — VME01 —', resumo.kwargs['text'])
            self.assertTrue(resumo.kwargs['text'].endswith(primeiro['descricao'].strip()))
            self.assertTrue(widgets('Item 1/2'))
            self.assertTrue(widgets('UNITÁRIO 1'))
            self.assertTrue(widgets('UNITÁRIO 2'))
            self.assertEqual(app.janela.maximizar_agendado, 250)
            self.assertFalse(widgets('ITEM 2'))
            for entry, valor in zip(campos(), ('8.000,00', '9.213,50')):
                entry.insert(0, valor)
            campos()[0].delete(0, 'end')
            campos()[0].insert(0, '8000')
            campos()[0].kwargs['<FocusOut>'](None)
            self.assertEqual(campos()[0].get(), '8.000,00')
            clicar('CONFIRMAR EMPRESA')
            self.assertTrue(widgets('ITEM 1'))  # valores novos exigem recálculo
            clicar('RECALCULAR')
            self.assertEqual(primeiro['referencias']['minimo'], 8000)
            retirar = widgets('RETIRAR VALOR — Fornecedor 2')[0]
            self.assertIn('Fornecedor 2', retirar.kwargs['text'])
            retirar.kwargs['variable'].set(True)
            clicar('RECALCULAR')
            self.assertEqual(primeiro['fornecedores_retirados'], ['F2'])
            historico = [w for w in widgets('Fornecedor 2') if w.kwargs.get('text_color') == '#FF5555']
            self.assertEqual(len(historico), 1)
            self.assertIn('12.990,00', historico[0].kwargs['text'])
            clicar('RETORNAR')
            self.assertEqual(len(primeiro['propostas']), 3)
            self.assertEqual(primeiro['fornecedores_retirados'], [])
            self.assertEqual(primeiro['precos_retirados'], [])
            self.assertFalse(widgets('RETORNAR'))
            self.salvar.assert_called()
            retirar = widgets('RETIRAR VALOR — Fornecedor 2')[0]
            retirar.kwargs['variable'].set(True)
            clicar('RECALCULAR')
            observacao = [w for w in Widget.criados if not w.destruido and w.kwargs.get('height') == 90][0]
            observacao.insert('1.0', 'Preço retirado após análise.')
            self.assertEqual(len(widgets('colocado —')), 2)
            clicar('CONFIRMAR EMPRESA')
            self.assertTrue(widgets('ITEM 2'))
            self.assertFalse(widgets('ITEM 1'))
            self.assertEqual([e.get() for e in campos()], ['', ''])
            self.assertEqual(primeiro['observacaoitem'], 'Preço retirado após análise.')
            clicar('VOLTAR')
            self.assertEqual([e.get() for e in campos()], ['8.000,00', '9.213,50'])
            self.assertEqual(len(widgets('colocado —')), 2)
            self.assertTrue(widgets('REFERÊNCIA:'))
            campos()[0].delete(0, 'end')
            campos()[0].insert(0, '8.500,00')
            clicar('RECALCULAR')
            clicar('CONFIRMAR EMPRESA')
            self.assertEqual(primeiro['pesquisas'][0], 8500)
            self.assertEqual(len(widgets('colocado —')), 7)
            for entry, valor in zip(campos(), ('1.000,00', '2.000,00')):
                entry.insert(0, valor)
            clicar('RECALCULAR')
            self.assertEqual(len([w for w in widgets('RETIRAR VALOR') if 'variable' in w.kwargs]), 7)
            self.assertEqual(segundo['pesquisas'], [1000, 2000])
            retirar = next(w for w in widgets('RETIRAR VALOR') if w.kwargs.get('text') == 'RETIRAR VALOR')
            retirar.kwargs['variable'].set(True)
            motivos = [w for w in Widget.criados if not w.destruido and
                       w.kwargs.get('placeholder_text') == 'Justificativa da retirada (obrigatória)']
            self.assertEqual(motivos[0].get(), 'Proposta enviada não atende ao descritivo do Termo de Referência conforme')
            self.assertIs(retirar.parent, motivos[0].parent)
            self.assertEqual(motivos[0].pack_config['side'], 'left')
            motivos[0].delete(0, 'end')
            clicar('RECALCULAR')
            self.assertEqual(len(segundo['propostas']), 7)
            motivos[0].insert(0, 'Proposta de item em desacordo com o solicitado no TR')
            clicar('RECALCULAR')
            self.assertEqual(segundo['fornecedores_retirados'], ['F0'])
            self.assertEqual(segundo['precos_retirados'][0]['justificativa'], motivos[0].get())
            self.assertAlmostEqual(segundo['referencias']['media'], (11+12+13+14+15+16+1000+2000)/8)
            self.assertTrue(widgets('FORNECEDORES RETIRADOS — ANÁLISE'))
            quinto = widgets('6º colocado')[0]
            quinto.kwargs['variable'].set(quinto.kwargs['value'])
            clicar('RECALCULAR')
            self.assertEqual(quinto.kwargs['variable'].get(), 5)
            clicar('CONFIRMAR EMPRESA')
            self.assertTrue(app.janela.destruido)
        app.assertions = exercitar
        with ExitStack() as stack:
            for nome in ('CTkToplevel', 'CTkFrame', 'CTkScrollableFrame', 'CTkLabel',
                         'CTkEntry', 'CTkTextbox', 'CTkCheckBox', 'CTkRadioButton', 'CTkButton'):
                stack.enter_context(patch.object(adj.ctk, nome, Widget))
            stack.enter_context(patch.object(adj.tk, 'IntVar', Var))
            stack.enter_context(patch.object(adj.tk, 'BooleanVar', Var))
            aviso = stack.enter_context(patch.object(adj.messagebox, 'showwarning'))
            erro = stack.enter_context(patch.object(adj.messagebox, 'showerror'))
            resultado = app._perguntar_pesquisas_precos([primeiro, segundo, cancelado])
        self.assertEqual(resultado, list(adj.TITULOS_PESQUISA_PADRAO))
        self.assertEqual(segundo['vencedor_idx'], 5)
        self.assertNotIn('referencias', cancelado)
        justificativa = adj.montar_texto_pesquisa_justificativa([primeiro, segundo, cancelado])
        self.assertTrue(justificativa.endswith('ITEM 1 Preço retirado após análise.'))
        self.assertEqual(justificativa.count('ITEM 1'), 1)
        self.assertEqual(aviso.call_count, 2)
        erro.assert_not_called()

    def test_retiradas_limites_independentes_sem_duplicar_historico(self):
        primeiro = item(precos=(1, 100, 1000), pesquisas=(100, 100))
        Widget.criados = []
        def vivos():
            return [w for w in Widget.criados if not w.destruido]
        class App(adj.AdjudicacaoOrdemMixin):
            def wait_window(app, janela):
                limites = [w for w in vivos() if w.kwargs.get('text', '').startswith('RETIRAR VALOR —')]
                self.assertEqual(len(limites), 2)
                for checkbox in limites:
                    checkbox.kwargs['variable'].set(True)
                # O mesmo fornecedor também pode ser marcado na classificação.
                checkbox = next(w for w in vivos() if w.kwargs.get('text') == 'RETIRAR VALOR')
                checkbox.kwargs['variable'].set(True)
                next(w for w in vivos() if w.kwargs.get('text') == 'RECALCULAR').kwargs['command']()
                self.assertEqual(len(primeiro['precos_retirados']), 2)
                self.assertEqual({p['grupo'] for p in primeiro['precos_retirados']}, {'superior', 'inferior'})
                self.assertEqual(primeiro['referencias']['media'], 100)
                for sufixo, fornecedor in [('ACIMA', 'Fornecedor 2'), ('ABAIXO', 'Fornecedor 0')]:
                    titulo = next(w for w in vivos() if w.kwargs.get('text') == 'FORNECEDORES RETIRADOS - 50% ' + sufixo)
                    registros = [w.kwargs.get('text', '') for linha in titulo.parent.children
                                 for w in [linha] + linha.children]
                    self.assertTrue(any(fornecedor in texto for texto in registros))
                    outro = 'Fornecedor 0' if sufixo == 'ACIMA' else 'Fornecedor 2'
                    self.assertFalse(any(outro in texto for texto in registros))
                next(w for w in vivos() if w.kwargs.get('text', '').startswith('CONFIRMAR EMPRESA')).kwargs['command']()
        with ExitStack() as stack:
            for nome in ('CTkToplevel', 'CTkFrame', 'CTkScrollableFrame', 'CTkLabel',
                         'CTkEntry', 'CTkTextbox', 'CTkCheckBox', 'CTkRadioButton', 'CTkButton'):
                stack.enter_context(patch.object(adj.ctk, nome, Widget))
            stack.enter_context(patch.object(adj.tk, 'IntVar', Var))
            stack.enter_context(patch.object(adj.tk, 'BooleanVar', Var))
            App()._perguntar_pesquisas_precos([primeiro])

    def test_observacao_na_justificativa_editavel(self):
        primeiro = item()
        primeiro['observacaoitem'] = 'Prazo confirmado.'
        texto = adj.montar_texto_pesquisa_justificativa([primeiro])
        Widget.criados = []
        class App(adj.AdjudicacaoOrdemMixin):
            def wait_window(self, janela):
                caixas = [w for w in Widget.criados if w.kwargs.get('height') == 350]
                self.teste.assertTrue(caixas[0].get().endswith('ITEM 1 Prazo confirmado.'))
                caixas[0].insert('1.0', texto + '\nComplemento editado.')
                botao = next(w for w in Widget.criados if w.kwargs.get('text') == 'Gerar Mapa Comparativo')
                botao.kwargs['command']()
        app = App()
        app.teste = self
        with ExitStack() as stack:
            for nome in ('CTkToplevel', 'CTkScrollableFrame', 'CTkLabel', 'CTkTextbox', 'CTkButton'):
                stack.enter_context(patch.object(adj.ctk, nome, Widget))
            resultado = app._perguntar_dados_finalizacao('', '', texto)
        self.assertTrue(resultado['texto_pesquisa'].endswith('ITEM 1 Prazo confirmado.\nComplemento editado.'))

    def test_todos_cancelados_sem_abrir_tela(self):
        it = item()
        it['cancelado'] = True
        with patch.object(adj.ctk, 'CTkToplevel') as janela:
            resultado = adj.AdjudicacaoOrdemMixin()._perguntar_pesquisas_precos([it])
        janela.assert_not_called()
        self.assertEqual(resultado, list(adj.TITULOS_PESQUISA_PADRAO))


class PersistenciaTelasTest(unittest.TestCase):
    def test_editor_relatorio_salva_reabre_e_atualiza_texto(self):
        import json
        import tempfile
        from pathlib import Path
        class App(adj.AdjudicacaoOrdemMixin):
            def _caminho_arquivo_mapa_empresas(self, *args):
                return self.arquivo
            def wait_window(self, janela):
                self.exercitar(janela)
        def caixa():
            return next(w for w in Widget.criados if not w.destruido and w.kwargs.get('wrap') == 'word')
        def clicar(texto):
            next(w for w in Widget.criados if not w.destruido and w.kwargs.get('text') == texto).kwargs['command']()
        with tempfile.TemporaryDirectory() as pasta, ExitStack() as stack:
            for nome in ('CTkToplevel', 'CTkFrame', 'CTkLabel', 'CTkTextbox', 'CTkButton'):
                stack.enter_context(patch.object(adj.ctk, nome, Widget))
            stack.enter_context(patch.object(adj.messagebox, 'askyesno', return_value=True))
            app = App()
            app._pasta_cotacao_atual, app._codigo_cotacao_atual = pasta, '178'
            app.arquivo = str(Path(pasta) / 'mapa.json')
            Path(app.arquivo).write_text(json.dumps({'decisoes_itens': {'1': {'cancelado': True}}}))
            def editar(janela):
                self.assertEqual(caixa().get(), 'Texto automático')
                caixa().insert('1.0', 'Texto editado com justificativa adicional')
                clicar('SALVAR TEXTO')
                janela.kwargs['WM_DELETE_WINDOW']()
            Widget.criados = []
            app.exercitar = editar
            self.assertIsNone(app._editar_relatorio_pesquisa('Texto automático'))
            salvo = json.loads(Path(app.arquivo).read_text())
            self.assertTrue(salvo['decisoes_itens']['1']['cancelado'])
            self.assertEqual(salvo['relatorio_pesquisa']['texto'], 'Texto editado com justificativa adicional')
            def reabrir(janela):
                self.assertEqual(caixa().get(), 'Texto editado com justificativa adicional')
                self.assertTrue(any('mudaram' in w.kwargs.get('text', '') for w in Widget.criados))
                clicar('ATUALIZAR COM DADOS ATUAIS')
                self.assertEqual(caixa().get(), 'Novos dados automáticos')
                clicar('SALVAR E GERAR DOCUMENTOS')
            Widget.criados = []
            app.exercitar = reabrir
            self.assertEqual(app._editar_relatorio_pesquisa('Novos dados automáticos'), 'Novos dados automáticos')
            salvo = json.loads(Path(app.arquivo).read_text())
            self.assertEqual(salvo['relatorio_pesquisa']['base'], 'Novos dados automáticos')

    def test_fechar_reabrir_editar_limpar_sem_apagar_outra_tela(self):
        import json
        import tempfile
        from pathlib import Path
        class App(adj.AdjudicacaoOrdemMixin):
            def _caminho_arquivo_mapa_empresas(self, *args):
                return self.arquivo
            def wait_window(self, janela):
                self.exercitar(janela)
        def widgets(texto):
            return [w for w in Widget.criados if not w.destruido and
                    texto in w.kwargs.get('text', '')]
        def clicar(texto):
            widgets(texto)[0].kwargs['command']()
        def valores():
            return [w for w in Widget.criados if not w.destruido and
                    w.kwargs.get('font') == ('Segoe UI', 20) and 'placeholder_text' not in w.kwargs
                    and w.kwargs.get('height') == 38]
        with tempfile.TemporaryDirectory() as pasta, ExitStack() as stack:
            for nome in ('CTkToplevel', 'CTkFrame', 'CTkScrollableFrame', 'CTkLabel',
                         'CTkEntry', 'CTkTextbox', 'CTkCheckBox', 'CTkRadioButton', 'CTkButton'):
                stack.enter_context(patch.object(adj.ctk, nome, Widget))
            stack.enter_context(patch.object(adj.tk, 'IntVar', Var))
            stack.enter_context(patch.object(adj.tk, 'BooleanVar', Var))
            stack.enter_context(patch.object(adj, 'centralizar_janela'))
            app = App()
            app._pasta_cotacao_atual, app._codigo_cotacao_atual = pasta, '1'
            app.arquivo = str(Path(pasta) / 'mapa.json')
            Path(app.arquivo).write_text(json.dumps({'fornecedores': {'original': {}},
                'decisoes_itens': {'1': {'cancelado': False}, '2': {'observacaoitem': 'Preservar'}}}))
            def dados():
                return json.loads(Path(app.arquivo).read_text())
            primeiro = item()
            def fechar(janela):
                valores()[0].delete(0, 'end')
                valores()[0].insert(0, '1234,56')
                janela.kwargs['WM_DELETE_WINDOW']()
            app.exercitar = fechar
            Widget.criados = []
            app._perguntar_pesquisas_precos([primeiro])
            self.assertEqual(dados()['decisoes_itens']['1']['pesquisas'][0], 1234.56)
            # Simula reconstrução dos itens após fechar e reabrir a adjudicação.
            segundo = item()
            segundo.update(dados()['decisoes_itens']['1'])
            def reabrir(janela):
                self.assertEqual(valores()[0].get(), '1.234,56')
                valores()[0].delete(0, 'end')
                valores()[0].insert(0, '2222')
                clicar('RECALCULAR')
                clicar('CONFIRMAR EMPRESA')
            app.exercitar = reabrir
            Widget.criados = []
            app._perguntar_pesquisas_precos([segundo])
            self.assertEqual(dados()['decisoes_itens']['1']['pesquisas'][0], 2222)
            segundo['cancelado'] = True
            app._salvar_campos_telas([segundo], ('cancelado',))
            segundo['cancelado'] = False  # permite exercitar a pesquisa isoladamente
            adj.aplicar_retiradas(segundo, ['F2'])
            def limpar_pesquisa(janela):
                clicar('LIMPAR DADOS')
                self.assertEqual([v.get() for v in valores()], ['', ''])
                self.assertEqual(len(segundo['propostas']), 3)
                self.assertEqual(segundo['fornecedores_retirados'], [])
                janela.kwargs['WM_DELETE_WINDOW']()
            app.exercitar = limpar_pesquisa
            Widget.criados = []
            app._perguntar_pesquisas_precos([segundo])
            self.assertTrue(dados()['decisoes_itens']['1']['cancelado'])
            self.assertEqual(dados()['decisoes_itens']['1']['pesquisas'], [None, None])
            segundo.update(dados()['decisoes_itens']['1'])
            def limpar_cancelados(janela):
                checkbox = widgets('Item 1')[0]
                self.assertTrue(checkbox.kwargs['variable'].get())
                self.assertEqual(checkbox.kwargs['font'], ('Segoe UI', 20))
                clicar('LIMPAR DADOS')
                self.assertFalse(checkbox.kwargs['variable'].get())
                janela.kwargs['WM_DELETE_WINDOW']()
            app.exercitar = limpar_cancelados
            Widget.criados = []
            app._perguntar_itens_cancelados([segundo])
            self.assertFalse(dados()['decisoes_itens']['1']['cancelado'])
            self.assertEqual(dados()['decisoes_itens']['1']['pesquisas'], [None, None])
            self.assertEqual(dados()['decisoes_itens']['2'], {'observacaoitem': 'Preservar'})
            self.assertEqual(dados()['fornecedores'], {'original': {}})


if __name__ == '__main__':
    unittest.main()
