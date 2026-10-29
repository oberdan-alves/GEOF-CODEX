import json
import tempfile
import unittest
from contextlib import ExitStack
from pathlib import Path
from unittest.mock import patch

import caminhos
import GeradorCotacao as principal


class ArmazenamentoCompartilhadoTest(unittest.TestCase):
    def test_executavel_em_pendrive_usa_recursos_do_pacote_e_destino_compartilhado(self):
        with tempfile.TemporaryDirectory() as pasta, ExitStack() as stack:
            usb = Path(pasta) / 'pendrive'
            pacote = Path(pasta) / 'extracao-exe'
            usb.mkdir()
            (pacote / 'assets').mkdir(parents=True)
            (pacote / 'assets' / 'brasao_gdf.png').write_bytes(b'logo-do-executavel')
            stack.enter_context(patch.object(caminhos, '_pasta_compartilhada', None))
            stack.enter_context(patch.object(caminhos.sys, 'frozen', True, create=True))
            stack.enter_context(patch.object(caminhos.sys, '_MEIPASS', str(pacote), create=True))
            stack.enter_context(patch.object(caminhos.sys, 'executable', str(usb / 'GeradorCotacao.exe')))
            destino = caminhos.configurar_pasta_compartilhada(str(Path(pasta) / 'drive' / 'GEOF'))
            caminhos.salvar_json_compartilhado('feriados.json', {'feriados': []})
            self.assertEqual((destino / 'assets' / 'brasao_gdf.png').read_bytes(), b'logo-do-executavel')
            self.assertFalse((usb / 'feriados.json').exists())
            self.assertFalse((pacote / 'feriados.json').exists())
            self.assertTrue((destino / 'feriados.json').exists())
            caminhos.desativar_pasta_compartilhada()
            with self.assertRaises(ValueError):
                caminhos.salvar_json_compartilhado('feriados.json', {})

    def test_deriva_pasta_windows_e_rede_unc(self):
        with patch.object(caminhos.os, 'name', 'nt'):
            self.assertEqual(caminhos.pasta_dados_compartilhados('G:/Meu Drive/GEOF'),
                'G:\\Meu Drive\\Planilhas Geradoras Cotação Pagamento Prestação\\GeradorCotacao')
            self.assertEqual(caminhos.pasta_dados_compartilhados('//servidor/compartilhamento/GEOF'),
                '\\\\servidor\\compartilhamento\\Planilhas Geradoras Cotação Pagamento Prestação\\GeradorCotacao')
            with self.assertRaises(ValueError):
                caminhos.pasta_dados_compartilhados('')

    def preparar(self, pasta, stack):
        local = Path(pasta) / 'pendrive'
        pacote = Path(pasta) / 'pacote-temporario'
        geof = Path(pasta) / 'drive-compartilhado' / 'GEOF'
        local.mkdir()
        (pacote / 'assets').mkdir(parents=True)
        (pacote / 'assets' / 'brasao_gdf.png').write_bytes(b'imagem-inicial')
        (local / 'config_padrao.json').write_text(json.dumps({'da_nome': 'Dados antigos', 'pasta_geof': str(geof)}))
        (local / 'feriados.json').write_text(json.dumps({'feriados': [{'data': '01/01/2027', 'feriado': 'Ano Novo'}]}))
        stack.enter_context(patch.object(caminhos, '_pasta_compartilhada', None))
        stack.enter_context(patch.object(caminhos, 'pasta_programa', return_value=local))
        stack.enter_context(patch.object(caminhos, '_pasta_recursos_empacotados', return_value=pacote))
        for nome in ('CONFIG_FILE', 'FERIADOS_FILE', 'CAMINHO_ASSETS', 'CAMINHO_LOGO_GDF'):
            stack.enter_context(patch.object(principal, nome, getattr(principal, nome)))
        return local, pacote, geof

    def test_migra_uma_vez_e_outro_executavel_le_mesmos_dados(self):
        with tempfile.TemporaryDirectory() as pasta, ExitStack() as stack:
            local, pacote, geof = self.preparar(pasta, stack)
            destino = principal.definir_armazenamento_compartilhado(str(geof))
            self.assertEqual((destino / 'assets' / 'brasao_gdf.png').read_bytes(), b'imagem-inicial')
            self.assertEqual(principal.carregar_config()['da_nome'], 'Dados antigos')
            config = principal.carregar_config()
            config.update(pasta_geof=str(geof), da_nome='Alteração do operador A')
            principal.salvar_config(config)
            principal.salvar_feriados([{'data': '25/12/2027', 'feriado': 'Natal compartilhado'}])
            (destino / 'assets' / 'brasao_gdf.png').write_bytes(b'imagem-compartilhada-editada')
            (local / 'config_padrao.json').write_text(json.dumps({'da_nome': 'Não deve substituir o compartilhado'}))
            caminhos._pasta_compartilhada = None  # Simula nova execução em outro local.
            principal.definir_armazenamento_compartilhado(str(geof))
            self.assertEqual(principal.carregar_config()['da_nome'], 'Alteração do operador A')
            self.assertEqual(principal.carregar_feriados()[0]['feriado'], 'Natal compartilhado')
            self.assertEqual(caminhos.caminho_recurso('assets/brasao_gdf.png'), destino / 'assets/brasao_gdf.png')
            self.assertEqual(caminhos.caminho_configuracao('config_padrao.json'), destino / 'config_padrao.json')
            self.assertEqual((destino / 'assets' / 'brasao_gdf.png').read_bytes(), b'imagem-compartilhada-editada')
            self.assertEqual(json.loads((local / 'config_padrao.json').read_text())['da_nome'], 'Não deve substituir o compartilhado')
            self.assertFalse(list(destino.glob('*.tmp')))

    def test_troca_de_raiz_recarrega_config_e_imagens_corretos(self):
        with tempfile.TemporaryDirectory() as pasta, ExitStack() as stack:
            _, _, geof = self.preparar(pasta, stack)
            primeiro = principal.definir_armazenamento_compartilhado(str(geof))
            caminhos.salvar_json_compartilhado('config_padrao.json', {'da_nome': 'Drive A'})
            outro_geof = Path(pasta) / 'outro-drive' / 'GEOF'
            segundo = principal.definir_armazenamento_compartilhado(str(outro_geof))
            caminhos.salvar_json_compartilhado('config_padrao.json', {'da_nome': 'Drive B'})
            self.assertNotEqual(primeiro, segundo)
            self.assertEqual(principal.carregar_config()['da_nome'], 'Drive B')
            self.assertTrue(principal.CAMINHO_LOGO_GDF.startswith(str(segundo)))
            principal.definir_armazenamento_compartilhado(str(geof))
            self.assertEqual(principal.carregar_config()['da_nome'], 'Drive A')

    def test_indisponibilidade_nao_salva_na_pasta_antiga_ou_no_executavel(self):
        with tempfile.TemporaryDirectory() as pasta, ExitStack() as stack:
            local, _, geof = self.preparar(pasta, stack)
            destino = principal.definir_armazenamento_compartilhado(str(geof))
            original = (destino / 'config_padrao.json').read_bytes()
            with patch.object(Path, 'mkdir', side_effect=PermissionError('Sem acesso à rede')):
                with self.assertRaises(PermissionError):
                    principal.definir_armazenamento_compartilhado(str(Path(pasta) / 'sem-acesso' / 'GEOF'))
            with self.assertRaises(ValueError):
                caminhos.salvar_json_compartilhado('config_padrao.json', {'da_nome': 'Não salvar'})
            self.assertEqual((destino / 'config_padrao.json').read_bytes(), original)
            self.assertNotEqual(principal.CONFIG_FILE, str(local / 'config_padrao.json'))


if __name__ == '__main__':
    unittest.main()
