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
    def get(self):
        return self.valor
    def geometry(self, tamanho):
        self.tamanho = tamanho
    def minsize(self, *args):
        pass
    def title(self, title):
        pass
    def grab_set(self):
        pass


class TelaPesquisaItemTest(unittest.TestCase):
    def test_recalculo_avanco_e_quinto_vencedor(self):
        primeiro = item()
        segundo = item(precos=(10, 11, 12, 13, 14), numero=2)
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
            self.assertEqual(len(widgets('colocado —')), 2)
            clicar('CONFIRMAR EMPRESA')
            self.assertTrue(widgets('ITEM 2'))
            self.assertFalse(widgets('ITEM 1'))
            self.assertEqual([e.get() for e in campos()], ['', ''])
            self.assertEqual(len(widgets('colocado —')), 5)
            quinto = widgets('5º colocado')[0]
            quinto.kwargs['variable'].set(quinto.kwargs['value'])
            clicar('RECALCULAR')
            self.assertEqual(quinto.kwargs['variable'].get(), 4)
            clicar('CONFIRMAR EMPRESA')
            self.assertTrue(app.janela.destruido)
        app.assertions = exercitar
        with ExitStack() as stack:
            for nome in ('CTkToplevel', 'CTkFrame', 'CTkScrollableFrame', 'CTkLabel',
                         'CTkEntry', 'CTkCheckBox', 'CTkRadioButton', 'CTkButton'):
                stack.enter_context(patch.object(adj.ctk, nome, Widget))
            stack.enter_context(patch.object(adj.tk, 'IntVar', Var))
            stack.enter_context(patch.object(adj.tk, 'BooleanVar', Var))
            aviso = stack.enter_context(patch.object(adj.messagebox, 'showwarning'))
            erro = stack.enter_context(patch.object(adj.messagebox, 'showerror'))
            resultado = app._perguntar_pesquisas_precos([primeiro, segundo, cancelado])
        self.assertEqual(resultado, list(adj.TITULOS_PESQUISA_PADRAO))
        self.assertEqual(segundo['vencedor_idx'], 4)
        self.assertNotIn('referencias', cancelado)
        aviso.assert_called_once()
        erro.assert_not_called()

    def test_todos_cancelados_sem_abrir_tela(self):
        it = item()
        it['cancelado'] = True
        with patch.object(adj.ctk, 'CTkToplevel') as janela:
            resultado = adj.AdjudicacaoOrdemMixin()._perguntar_pesquisas_precos([it])
        janela.assert_not_called()
        self.assertEqual(resultado, list(adj.TITULOS_PESQUISA_PADRAO))


if __name__ == '__main__':
    unittest.main()
