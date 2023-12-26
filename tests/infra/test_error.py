import os
import unittest
from unittest.mock import MagicMock, patch

from realce.infra.error import existe_erro, tratar_pasta_destino, is_valid_file


class ErrorTestCase(unittest.TestCase):
    def setUp(self):
        self.campo_arquivo_excel = MagicMock()
        self.campo_arquivo_pdf = MagicMock()
        self.pasta_destino = MagicMock()

    @patch('realce.infra.error.messagebox.showerror')
    def test_campos_excel_e_pdf_estao_vazios(self, mock_showerror):
        self.campo_arquivo_excel.text.return_value = ''
        self.campo_arquivo_pdf.text.return_value = ''
        result_existe_erro = existe_erro(self.campo_arquivo_excel.text(), self.campo_arquivo_pdf.text())
        mock_showerror.assert_called_once()
        self.assertIn('Por favor, selecione o arquivo Excel e o arquivo PDF.', mock_showerror.call_args.args[1])
        self.assertTrue(result_existe_erro)

    @patch('realce.infra.error.messagebox.showerror')
    def test_pasta_destino_nao_foi_selecionada(self, mock_showerror):
        self.pasta_destino = None
        result_pasta_destino = tratar_pasta_destino(self.pasta_destino)
        mock_showerror.assert_called_once()
        self.assertIn('Selecione uma pasta de destino válida.', mock_showerror.call_args.args[1])
        self.assertTrue(result_pasta_destino)

    def test_pasta_destino_foi_selecionada_corretamente(self):
        self.pasta_destino = os.getcwd()
        result_pasta_destino = tratar_pasta_destino(self.pasta_destino)
        self.assertFalse(result_pasta_destino)

    @patch('realce.infra.error.messagebox.showerror')
    def test_extensao_do_arquivo_nao_eh_suportada(self, mock_showerror):
        self.campo_arquivo_excel.text.return_value = 'extensao_de_planilha_nao_suportada.docx'
        self.campo_arquivo_pdf.text.return_value = 'extensao_de_pdf_nao_suportada.docx'
        result_arquivo_invalido = is_valid_file(self.campo_arquivo_excel.text(), self.campo_arquivo_pdf.text())
        mock_showerror.assert_called_once()
        self.assertIn('Ocorreu um problema', mock_showerror.call_args.args[0])
        self.assertFalse(result_arquivo_invalido)

    def test_extensao_do_arquivo_eh_valida(self):
        self.campo_arquivo_excel.text.return_value = 'extensao_de_planilha_valida.xlsx'
        self.campo_arquivo_pdf.text.return_value = 'extensao_pdf_valida.pdf'
        result_arquivo_valido = is_valid_file(self.campo_arquivo_excel.text(), self.campo_arquivo_pdf.text())
        self.assertTrue(result_arquivo_valido)
