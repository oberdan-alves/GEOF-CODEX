import unittest
from unittest.mock import Mock
from contextlib import ExitStack
from unittest.mock import patch

import Mapa
from test_tela_pesquisa_item import Widget


class Campo(Widget):
    def grid(self, **kwargs):
        self.grid_config = kwargs
    def grid_columnconfigure(self, *args, **kwargs):
        pass


class CadastroPropostasTest(unittest.TestCase):
    def test_formatos(self):
        self.assertEqual(Mapa.formatar_quantidade_proposta(1.0), '1')
        self.assertEqual(Mapa.formatar_quantidade_proposta(1.5), '1,5')
        self.assertEqual(Mapa.formatar_unitario_proposta(1.0), '1,0000')
        self.assertEqual(Mapa.formatar_unitario_proposta('1234,5678'), '1234,5678')
        self.assertEqual(Mapa.formatar_unitario_proposta('1,23456'), '1,2346')
        self.assertEqual(Mapa.numero_proposta('1.234,5678'), 1234.5678)
        self.assertIsNone(Mapa.numero_proposta('nan'))

    def criar_formulario(self):
        app = Mapa.MapaMixin()
        app.wizard_conteudo, app.wizard_rodape, app.wizard_cabecalho = Campo(), Campo(), Campo()
        app.lista_empresas_mapa = ['EMPRESA']
        app.indice_fornecedor_atual = 0
        app._dados_form_atual = None
        app._blink_jobs, app.saldos_bd_atual = {}, {}
        app.mapa_fornecedores = {}
        app.linhas_resultado = [dict(ITEM=float(i), CODIGO=str(i), DESCRICAO='Serviço', UN='un', QTD=1.0)
                               for i in (1, 2, 3)]
        with ExitStack() as stack:
            for nome in ('CTkFrame', 'CTkScrollableFrame', 'CTkLabel', 'CTkEntry', 'CTkButton'):
                stack.enter_context(patch.object(Mapa.ctk, nome, Campo))
            app._renderizar_form_fornecedor()
        return app

    @staticmethod
    def escrever(campo, valor):
        campo.delete(0, 'end')
        campo.insert(0, valor)

    def test_destaque_acompanha_edicao_e_restaura_cores_alternadas(self):
        app = self.criar_formulario()
        primeiro, segundo, terceiro = app._entradas_item_atual
        self.assertEqual(primeiro['faixa'].kwargs['fg_color'], terceiro['faixa'].kwargs['fg_color'])
        self.assertNotEqual(primeiro['faixa'].kwargs['fg_color'], segundo['faixa'].kwargs['fg_color'])
        primeiro['entry_marca'].kwargs['<FocusIn>'](None)
        self.assertEqual(primeiro['faixa'].kwargs['fg_color'], '#e2e2e2')
        self.assertEqual(primeiro['faixa'].kwargs['border_color'], '#8c8c8c')
        self.assertTrue(all(label.kwargs['fg_color'] == '#e2e2e2' for label in primeiro['labels']))
        segundo['entry_unit'].kwargs['<FocusIn>'](None)
        self.assertEqual(primeiro['faixa'].kwargs['fg_color'], primeiro['fundo'])
        self.escrever(segundo['entry_unit'], '1.23456')
        segundo['entry_unit'].kwargs['<FocusOut>'](None)
        self.assertEqual(segundo['entry_unit'].get(), '1,2346')
        self.assertEqual(segundo['faixa'].kwargs['fg_color'], segundo['fundo'])
        self.assertEqual(segundo['label_total'].kwargs['text_color'], segundo['texto'])

    def test_repetir_e_desfazer_preserva_campos_vazios_e_data(self):
        app = self.criar_formulario()
        primeiro, segundo, terceiro = app._entradas_item_atual
        self.escrever(app.wizard_entry_data, '14/10/2026')
        for chave, valor in dict(entry_qtd='3', entry_unit='1,2345', entry_prazo='10').items():
            self.escrever(primeiro[chave], valor)
        self.escrever(segundo['entry_marca'], 'Marca original')
        self.escrever(segundo['entry_unit'], '8,7654')
        app._repetir_dados_primeiro_item()
        self.assertEqual(segundo['entry_qtd'].get(), '3')
        self.assertEqual(terceiro['entry_unit'].get(), '1,2345')
        self.assertEqual(segundo['entry_marca'].get(), 'Marca original')
        self.assertEqual(segundo['entry_prazo'].get(), '10')
        self.assertEqual(segundo['label_total'].kwargs['text'], 'R$ 3,70')
        self.assertEqual(app.wizard_entry_data.get(), '14/10/2026')
        app._desfazer_repeticao()
        self.assertEqual(segundo['entry_unit'].get(), '8,7654')
        self.assertEqual(terceiro['entry_unit'].get(), '')
        self.assertEqual(segundo['entry_qtd'].get(), '1')
        self.assertEqual(primeiro['entry_unit'].get(), '1,2345')
        self.assertEqual(app.wizard_botao_desfazer.kwargs['state'], 'disabled')

    def test_mesma_grade_e_persistencia_quatro_casas(self):
        Widget.criados = []
        app = self.criar_formulario()
        registro = app._entradas_item_atual[0]
        self.assertEqual(registro['entry_qtd'].get(), '1')
        cabecalho = next(w for w in Widget.criados if w.kwargs.get('text') == 'UNITÁRIO')
        self.assertIs(cabecalho.parent, registro['entry_unit'].parent)
        self.assertEqual(cabecalho.grid_config['column'], registro['entry_unit'].grid_config['column'])
        labels_item = [w for w in Widget.criados if w.kwargs.get('text') == '1']
        self.assertGreaterEqual(len(labels_item), 2)  # ITEM e QTD sem .0
        self.escrever(app.wizard_entry_data, '14/10/2026')
        self.escrever(registro['entry_unit'], '12.3456')
        app._renderizar_preview_fornecedor = Mock()
        app._salvar_formulario_atual()
        self.assertEqual(registro['entry_unit'].get(), '12,3456')
        self.assertEqual(app._dados_form_atual['itens'][0]['unitario'], 12.3456)
        self.assertEqual(app._dados_form_atual['itens'][0]['total'], 12.35)
        app._renderizar_preview_fornecedor.assert_called_once()
