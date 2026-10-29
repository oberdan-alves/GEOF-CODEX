import unittest
from contextlib import ExitStack
from unittest.mock import patch

import Emails
from test_tela_pesquisa_item import Widget, Var


class EmailCotacaoEditavelTest(unittest.TestCase):
    def test_tela_exibe_texto_atualizado_e_utiliza_edicoes(self):
        class App(Emails.EmailMixin):
            def wait_window(self, janela):
                campo = next(w for w in Widget.criados if w.kwargs.get('wrap') == 'word')
                for trecho in ('Medicamento informar o GGREM e Nº Anvisa.', '13.6 - Para venda de MEDICAMENTOS', '14- O contratado deverá manter'):
                    self.teste.assertIn(trecho, campo.get())
                campo.insert('1.0', 'Texto editado pelo operador.\nCondição específica.')
                hora = next(w for w in Widget.criados if w.kwargs.get('placeholder_text') == '18:00')
                hora.insert(0, '18:00')
                next(w for w in Widget.criados if w.kwargs.get('text') == 'Gerar e-mail').kwargs['command']()
        with ExitStack() as stack:
            for nome in ('CTkToplevel', 'CTkFrame', 'CTkScrollableFrame', 'CTkLabel', 'CTkOptionMenu',
                         'CTkEntry', 'CTkCheckBox', 'CTkTextbox', 'CTkButton'):
                stack.enter_context(patch.object(Emails.ctk, nome, Widget))
            for nome in ('StringVar', 'BooleanVar'):
                stack.enter_context(patch.object(Emails.tk, nome, Var))
            stack.enter_context(patch.object(Emails, 'centralizar_janela'))
            Widget.criados = []
            app = App()
            app.teste = self
            dados = app._perguntar_dados_email_cotacao()
        assunto, corpo, indice = app._montar_texto_email_cotacao(['123'], '2026', [], dados['publicacao'],
            dados['horas'], '01/11/2026', dados['exclusividade'], texto_corpo=dados['texto_corpo'])
        self.assertIn('Texto editado pelo operador.\nCondição específica.', corpo)
        self.assertNotIn('13.6 -', corpo)
        self.assertTrue(corpo.splitlines()[indice].startswith('2 - Prazo'))
        self.assertIn('123/2026', assunto)


if __name__ == '__main__':
    unittest.main()
