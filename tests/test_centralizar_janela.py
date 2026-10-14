import unittest
from GeradorCotacao import centralizar_janela


class Janela:
    def __init__(self, tamanho, tela, escala=1):
        self.tamanho, self.tela, self.escala = tamanho, tela, escala
        self.minimo = (950, 750)
    def after_idle(self, callback):
        self.callback = callback
    def update_idletasks(self):
        pass
    def geometry(self, valor=None):
        if valor is not None:
            self.tamanho = valor
        return self.tamanho
    def winfo_screenwidth(self):
        return self.tela[0]
    def winfo_screenheight(self):
        return self.tela[1]
    def minsize(self, *valores):
        if valores:
            self.minimo = valores
        return self.minimo
    def _get_window_scaling(self):
        return self.escala


class CentralizarTest(unittest.TestCase):
    def test_posiciona_no_centro(self):
        janela = Janela('1200x950-50-20', (1920, 1080))
        centralizar_janela(janela)
        janela.callback()
        self.assertEqual(janela.tamanho, '1200x950+360+65')

    def test_tela_menor_com_escala(self):
        janela = Janela('1280x940', (1366, 768), 1.25)
        centralizar_janela(janela)
        janela.callback()
        self.assertEqual(janela.tamanho, '1052x534+25+50')
        self.assertEqual(janela.minimo, (950, 534))


class MaximizarTest(unittest.TestCase):
    def test_maximiza_apos_criar_janela(self):
        from unittest.mock import Mock
        from GeradorCotacao import maximizar_janela
        janela = Mock()
        maximizar_janela(janela)
        janela.after.assert_called_once()
        janela.after.call_args.args[1]()
        janela.state.assert_called_once_with('zoomed')

    def test_linux_usa_atributo_zoomed(self):
        from unittest.mock import Mock
        import tkinter as tk
        from GeradorCotacao import maximizar_janela
        janela = Mock()
        janela.state.side_effect = tk.TclError('zoomed indisponível')
        maximizar_janela(janela)
        janela.after.call_args.args[1]()
        janela.attributes.assert_called_once_with('-zoomed', True)
