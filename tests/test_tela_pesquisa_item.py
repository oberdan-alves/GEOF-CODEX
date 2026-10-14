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
        pass
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
    def protocol(self, *args):
        pass
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
            self.assertTrue(resumo.kwargs['text'].endswith('...'))
            self.assertEqual(app.janela.maximizar_agendado, 250)
            self.assertFalse(widgets('ITEM 2'))
            for entry, valor in zip(campos(), ('8.000,00', '9.213,50')):
                entry.insert(0, valor)
            clicar('CONFIRMAR EMPRESA')
            self.assertTrue(widgets('ITEM 1'))  # valores novos exigem recálculo
            clicar('RECALCULAR')
            self.assertEqual(primeiro['referencias']['minimo'], 8000)
            retirar = widgets('RETIRAR VALOR')[0]
            self.assertIn('Fornecedor 2', retirar.kwargs['text'])
            retirar.kwargs['variable'].set(True)
            clicar('RECALCULAR')
            self.assertEqual(primeiro['fornecedores_retirados'], ['F2'])
            historico = [w for w in widgets('Fornecedor 2') if w.kwargs.get('text_color') == '#FF5555']
            self.assertEqual(len(historico), 1)
            self.assertIn('12.990,00', historico[0].kwargs['text'])
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
            self.assertEqual(widgets('RETIRAR VALOR'), [])
            self.assertEqual(segundo['pesquisas'], [1000, 2000])
            quinto = widgets('7º colocado')[0]
            quinto.kwargs['variable'].set(quinto.kwargs['value'])
            clicar('RECALCULAR')
            self.assertEqual(quinto.kwargs['variable'].get(), 6)
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
        self.assertEqual(segundo['vencedor_idx'], 6)
        self.assertNotIn('referencias', cancelado)
        justificativa = adj.montar_texto_pesquisa_justificativa([primeiro, segundo, cancelado])
        self.assertTrue(justificativa.endswith('ITEM 1 Preço retirado após análise.'))
        self.assertEqual(justificativa.count('ITEM 1'), 1)
        aviso.assert_called_once()
        erro.assert_not_called()

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


if __name__ == '__main__':
    unittest.main()
