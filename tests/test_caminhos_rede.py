import ntpath
import unittest
from types import SimpleNamespace
from unittest.mock import Mock, patch

import caminhos
import Cotacao
import Mapa


class CaminhosRedeTest(unittest.TestCase):
    def test_normalizacao_windows_unc_unidade_aspas_e_vazio(self):
        casos = {
            '//srv-fs/hrg_geof/GEOF': r'\\srv-fs\hrg_geof\GEOF',
            '//srv-fs/hrg_geof/GEOF\\HRG/PDPAS 2026/COTACAO/255/PROPOSTAS':
                r'\\srv-fs\hrg_geof\GEOF\HRG\PDPAS 2026\COTACAO\255\PROPOSTAS',
            r'\\srv-fs\hrg_geof\GEOF' + '\\': r'\\srv-fs\hrg_geof\GEOF',
            ' G:/Meu Drive/GEOF/ ': r'G:\Meu Drive\GEOF',
            '"//srv-fs/hrg_geof/GEOF"': r'\\srv-fs\hrg_geof\GEOF',
            '': '', '   ': '', None: '',
        }
        with patch.object(caminhos.os, 'name', 'nt'):
            for entrada, esperado in casos.items():
                with self.subTest(entrada=entrada):
                    self.assertEqual(caminhos.normalizar_caminho(entrada), esperado)

    def test_acessar_propostas_usa_raiz_atual_e_mesma_pasta_do_pedido(self):
        class App(Cotacao.CotacaoMixin, Mapa.MapaMixin):
            pass
        app = App()
        app.entry_pasta_geof = Mock()
        app.entry_ano = Mock(get=Mock(return_value='2026'))
        app._sufixo_planilha_atual = 'HRG'
        app.indice_cotacao_ativa = 0
        app.entradas_cotacao = [Mock(get=Mock(return_value='255'))]
        criar, abrir = Mock(), Mock()
        with patch.object(caminhos.os, 'name', 'nt'), \
             patch.object(Cotacao, 'os', SimpleNamespace(path=ntpath, makedirs=criar)), \
             patch.object(Mapa, 'os', SimpleNamespace(startfile=abrir)), \
             patch.object(Mapa.messagebox, 'showerror') as erro:
            for raiz in ('G:/Meu Drive/GEOF', '//srv-fs/hrg_geof/GEOF', '//outro/compartilhamento/GEOF'):
                app.entry_pasta_geof.get.return_value = raiz
                pasta_pedido, _, _ = app._obter_pasta_cotacao('255')
                app.acao_acessar_propostas()
                esperado = ntpath.join(ntpath.normpath(raiz), 'HRG', 'PDPAS 2026', 'COTACAO', '255', 'PROPOSTAS')
                abrir.assert_called_with(esperado)
                criar.assert_called_with(esperado, exist_ok=True)
                self.assertEqual(ntpath.dirname(esperado), pasta_pedido)
            erro.assert_not_called()

    def test_caminho_posix_preservado(self):
        with patch.object(caminhos.os, 'name', 'posix'):
            self.assertEqual(caminhos.normalizar_caminho('/mnt/rede/GEOF/'), '/mnt/rede/GEOF')
            self.assertEqual(caminhos.normalizar_caminho(''), '')


if __name__ == '__main__':
    unittest.main()
