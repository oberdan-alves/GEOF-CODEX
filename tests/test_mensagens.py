import unittest
from unittest.mock import patch, Mock
from Mensagens import messagebox


class MensagensTest(unittest.TestCase):
    def test_confirmacoes_preservam_retornos(self):
        for metodo, resposta, esperado in [
            ('askyesno', 'yes', True), ('askyesno', 'no', False),
            ('askyesnocancel', 'cancel', None), ('askyesnocancel', 'no', False),
            ('askyesnocancel', 'yes', True), ('askokcancel', 'cancel', False),
            ('askokcancel', 'ok', True), ('askretrycancel', 'retry', True),
            ('askretrycancel', 'cancel', False), ('askquestion', 'no', 'no'),
        ]:
            with self.subTest(metodo=metodo, resposta=resposta):
                with patch.object(messagebox, '_show', return_value=resposta):
                    self.assertEqual(getattr(messagebox, metodo)('Título', 'Mensagem'), esperado)

    def test_utiliza_parent_e_opcoes_sem_criar_raiz(self):
        parent = Mock()
        with patch('Mensagens._Mensagem', return_value=Mock(result='yes')) as dialogo:
            self.assertTrue(messagebox.askyesno('Confirmar', 'Mensagem', parent=parent,
                                              default='no', detail='Detalhes'))
            dialogo.assert_called_once_with(parent, 'Confirmar', 'Mensagem', 'yesno',
                                            'no', 'Detalhes', 'question')


if __name__ == '__main__':
    unittest.main()
