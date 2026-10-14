import tempfile
import unittest
from pathlib import Path
from unittest.mock import Mock, patch
import GeradorCotacao as principal


class ConsultarMatrixTest(unittest.TestCase):
    def app(self, pasta, tipo='HRG'):
        app = principal._AppCotacaoBase.__new__(principal._AppCotacaoBase)
        app.entry_pasta_geof = Mock(get=Mock(return_value=str(pasta)))
        app.entry_ano = Mock(get=Mock(return_value='2026'))
        app.var_nome_planilha = Mock(get=Mock(return_value=f'Matrix_2026_{tipo}'))
        app._sufixo_planilha_atual = tipo
        app.entry_caminho_completo = Mock()
        app._salvar_defaults_atuais = Mock()
        app._abrir_planilha_matrix = Mock()
        return app

    def test_abre_planilha_selecionada_e_atualiza_caminho(self):
        with tempfile.TemporaryDirectory() as pasta:
            for tipo in ['APS', 'HRG', 'APS INV', 'HRG INV']:
                with self.subTest(tipo=tipo):
                    arquivo = Path(pasta)/tipo/'PDPAS 2026'/'MATRIX'/f'Matrix_2026_{tipo}.xlsm'
                    arquivo.parent.mkdir(parents=True)
                    arquivo.write_bytes(b'modelo preservado')
                    app = self.app(pasta, tipo)
                    app.acao_consultar_matrix()
                    app._abrir_planilha_matrix.assert_called_once_with(str(arquivo))
                    self.assertEqual(app.caminho_arquivo_atual, str(arquivo))
                    self.assertEqual(arquivo.read_bytes(), b'modelo preservado')
                    app.entry_caminho_completo.configure.assert_called_with(state='readonly')

    def test_arquivo_ausente_avisa_sem_abrir(self):
        with tempfile.TemporaryDirectory() as pasta:
            app = self.app(pasta)
            with patch.object(principal.messagebox, 'showerror') as aviso:
                app.acao_consultar_matrix()
            aviso.assert_called_once()
            app._abrir_planilha_matrix.assert_not_called()
            self.assertIsNone(app.caminho_arquivo_atual)

    def test_falha_de_abertura_exibe_erro(self):
        app = self.app('/GEOF')
        app._localizar_arquivo_matrix = Mock(return_value='/GEOF/Matrix.xlsx')
        app._abrir_planilha_matrix.side_effect = OSError('Excel indisponível')
        with patch.object(principal.messagebox, 'showerror') as aviso:
            app.acao_consultar_matrix()
        self.assertIn('Excel indisponível', aviso.call_args.args[1])

    def test_campos_invalidos_nao_abrem(self):
        for campo, valor in [('entry_pasta_geof', ''), ('entry_ano', 'abcd'), ('var_nome_planilha', '')]:
            app = self.app('/GEOF')
            getattr(app, campo).get.return_value = valor
            with patch.object(principal.messagebox, 'showwarning') as aviso:
                app.acao_consultar_matrix()
            aviso.assert_called_once()
            app._abrir_planilha_matrix.assert_not_called()

    def test_windows_abre_caminho_com_espacos(self):
        with tempfile.TemporaryDirectory() as pasta:
            arquivo = Path(pasta)/'Matrix_2026_HRG INV.xlsm'
            arquivo.touch()
            with patch.object(principal.sys, 'platform', 'win32'), \
                 patch.object(principal.os, 'startfile', create=True) as abrir:
                principal._AppCotacaoBase._abrir_planilha_matrix(str(arquivo))
            abrir.assert_called_once_with(str(arquivo.resolve()))
