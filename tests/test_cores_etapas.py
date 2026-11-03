import json
import tempfile
import unittest
from pathlib import Path
from unittest.mock import Mock, patch
from types import SimpleNamespace

from GeradorCotacao import _AppCotacaoBase
import Mapa


class CoresEtapasTest(unittest.TestCase):
    def setUp(self):
        self.temp = tempfile.TemporaryDirectory()
        self.addCleanup(self.temp.cleanup)
        self.app = _AppCotacaoBase.__new__(_AppCotacaoBase)
        self.app.indice_cotacao_ativa = 0
        self.app.entradas_cotacao = [Mock(get=Mock(return_value='257'))]
        self.app.entry_pasta_geof = Mock(get=Mock(return_value=self.temp.name))
        self.app.entry_ano = Mock(get=Mock(return_value='2026'))
        self.app._sufixo_planilha_atual = 'HRG'
        self.app._botoes_etapas = {n: Mock() for n in (
            'CRIAR PASTA COTAÇÃO', 'PEDIDO', 'LISTAR EMPRESAS',
            'INCLUIR EMPRESAS MAPA', 'ADJUDICAÇÃO', 'ORDEM')}
        self.pasta = Path(self.temp.name) / 'HRG' / 'PDPAS 2026' / 'COTACAO' / '257'
        self.pasta.mkdir(parents=True)

    def verificar(self, concluidos):
        self.app._atualizar_cores_etapas()
        for nome, botao in self.app._botoes_etapas.items():
            feito = nome in concluidos
            botao.configure.assert_called_with(
                fg_color='#333333' if feito else '#1a1a1a',
                text_color='#90ee90' if feito else '#ffffff')

    def salvar(self, nome, dados):
        (self.pasta / nome).write_text(json.dumps(dados), encoding='utf-8')

    def test_reconhece_resultados_existentes_sem_executar_acoes(self):
        (self.pasta / 'PROPOSTAS').mkdir()
        for nome in ('PEDIDO DE COTAÇÃO 257 HRG(NECFM).pdf',
                     'MAPA COMPARATIVO 257.pdf', 'ORDEM 257 EMPRESA.pdf'):
            (self.pasta / nome).touch()
        self.salvar('MAPA_EMPRESAS_257.json', {'fornecedores': {'A': {}}})
        self.verificar(set(self.app._botoes_etapas))

    def test_listagem_persistida_e_cadastro_parcial(self):
        self.salvar('EMPRESAS_LISTADAS_257.json', ['A', 'B'])
        self.salvar('MAPA_EMPRESAS_257.json', {'fornecedores': {'A': {}}})
        self.verificar({'LISTAR EMPRESAS'})
        self.salvar('MAPA_EMPRESAS_257.json', {'fornecedores': {'A': {}, 'B': {}}})
        self.verificar({'LISTAR EMPRESAS', 'INCLUIR EMPRESAS MAPA'})

    def test_isola_codigo_ano_tipo_e_limpeza(self):
        self.salvar('MAPA_EMPRESAS_257.json', {'fornecedores': {'A': {}}})
        self.verificar({'LISTAR EMPRESAS', 'INCLUIR EMPRESAS MAPA'})
        for campo in (self.app.entradas_cotacao[0], self.app.entry_ano):
            anterior = campo.get.return_value
            campo.get.return_value = '258' if anterior == '257' else '2025'
            self.verificar(set())
            campo.get.return_value = anterior
        self.app._sufixo_planilha_atual = 'APS'
        self.verificar(set())
        self.app._sufixo_planilha_atual = 'HRG'
        self.app.indice_cotacao_ativa = None
        self.verificar(set())

    def test_nao_confunde_previo_ou_arquivos_de_outro_codigo(self):
        for nome in ('MAPA COMPARATIVO PRÉVIO RESUMIDO 257.pdf',
                     'PEDIDO DE COTAÇÃO 258 HRG.pdf', 'ORDEM 258 A.pdf'):
            (self.pasta / nome).touch()
        (self.pasta / 'MAPA_EMPRESAS_257.json').write_text('{invalido')
        self.salvar('EMPRESAS_LISTADAS_257.json', [{'invalido': True}])
        self.verificar(set())

    def test_dica_acompanha_estado_da_cor_e_troca_de_cotacao(self):
        texto = 'PEDIDO'
        self.verificar(set())
        descricao = self.app._texto_dica_botao(texto)
        self.assertIn('Gera o Pedido', descricao)
        self.assertNotIn('já executada', descricao)
        (self.pasta / 'PEDIDO DE COTAÇÃO 257 HRG.pdf').touch()
        self.verificar({'PEDIDO'})
        self.assertIn('Ação já executada', self.app._texto_dica_botao(texto))
        self.assertIn(descricao, self.app._texto_dica_botao(texto))
        self.app.entradas_cotacao[0].get.return_value = '258'
        self.verificar(set())
        self.assertEqual(self.app._texto_dica_botao(texto), descricao)
        self.assertEqual(self.app._texto_dica_botao('Copiar', 'Copia o SIS.'), 'Copia o SIS.')

    def test_dica_cancela_espera_ao_sair_e_apos_clique(self):
        from GeradorCotacao import DicaBotao
        botao = Mock()
        dica = DicaBotao(botao, lambda: 'Ação')
        eventos = {c.args[0]: c.args[1] for c in botao.bind.call_args_list}
        eventos['<Enter>']()
        botao.after.assert_called_once_with(400, dica.mostrar)
        eventos['<Leave>']()
        botao.after_cancel.assert_called_once_with(botao.after.return_value)
        self.assertIsNone(dica.agendamento)
        dica.janela = Mock()
        janela = dica.janela
        eventos['<ButtonPress-1>']()
        janela.destroy.assert_called_once_with()
        self.assertIsNone(dica.janela)
        eventos['<Destroy>']()
        janela.destroy.assert_called_once_with()

    def test_comando_permanece_executavel_e_retorna_resultado(self):
        comando = Mock(return_value='resultado')
        self.assertEqual(self.app._executar_e_atualizar_etapas(comando), 'resultado')
        comando.assert_called_once_with()

    def test_listar_empresas_persiste_somente_listagem_resolvida(self):
        app = SimpleNamespace(
            entry_pasta_geof=self.app.entry_pasta_geof, entry_ano=self.app.entry_ano,
            indice_cotacao_ativa=0, entradas_cotacao=self.app.entradas_cotacao,
            _sufixo_planilha_atual='HRG', _obter_caminho_arquivo=Mock(return_value='matrix.xlsx'),
            _atualizar_cores_etapas=Mock(), _mostrar_resultado_empresas=Mock())
        with patch.object(Mapa, 'listar_arquivos_propostas', return_value=['A']), \
             patch.object(Mapa, 'carregar_fornecedores', return_value=[{
                 'nome': 'A', 'nome_normalizado': 'A'}]):
            Mapa.MapaMixin.acao_listar_empresas(app)
        registro = self.pasta / 'EMPRESAS_LISTADAS_257.json'
        self.assertEqual(json.loads(registro.read_text(encoding='utf-8')), ['A'])
        self.verificar({'LISTAR EMPRESAS'})
        app._atualizar_cores_etapas.assert_called_once_with()


if __name__ == '__main__':
    unittest.main()
