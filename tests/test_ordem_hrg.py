import tempfile
import unittest
from pathlib import Path
import fitz
from AdjudicacaoOrdem import gerar_pdf_ordem_compra


class OrdemHRGTest(unittest.TestCase):
    def texto_ordem(self, tipo='HRG', contatos=None):
        with tempfile.TemporaryDirectory() as pasta:
            arquivo = Path(pasta)/'ordem.pdf'
            gerar_pdf_ordem_compra(str(arquivo), '178', '2026', '033/2026', 'NFH',
                {'nome': 'Fornecedor teste'}, [], tipo, '', '', '', contatos=contatos)
            with fitz.open(arquivo) as pdf:
                return ' '.join(' '.join(p.get_text() for p in pdf).split())

    def test_hrg_tem_tres_contatos_da_imagem(self):
        texto = self.texto_ordem()
        for trecho in ['Informações técnicas/Aceite NF:', 'Farmácia NFH', 'MIRELLE CARDARELLI',
                       '3449-7070/7071/7073', 'Almoxarifado NUAL', 'WEILLER MACHADO NOBRE ALVES',
                       '3449-7092/7093/7094', 'Engenharia Clínica NECFM', 'RODRIGO TENÓRIO',
                       'Mezanino Acima do PAI no HRG', '9639-0162', 'MANUTENÇÃO EQUIPAMENTOS']:
            self.assertIn(trecho, texto)

    def test_contatos_customizados_preservados(self):
        texto = self.texto_ordem(contatos={'NFH': 'Farmacêutica atual', 'NECFM_HRG': 'Técnico atual'})
        self.assertIn('Farmacêutica atual', texto)
        self.assertIn('Técnico atual', texto)
        self.assertIn('WEILLER MACHADO', texto)

    def test_hrg_inv_mantem_entrega_npdoc(self):
        texto = self.texto_ordem('HRG INV')
        self.assertIn('Chefe do NPDOC', texto)
        self.assertNotIn('Engenharia Clínica NECFM', texto)
