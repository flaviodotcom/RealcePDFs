import unittest
from unittest.mock import MagicMock, patch

from realce.infra.error import existe_erro, is_valid_file, campos_sao_validos


class ErrorTestCase(unittest.TestCase):
    def setUp(self):
        self.campo_arquivo_excel = MagicMock()
        self.campo_arquivo_pdf = MagicMock()

    @patch('realce.infra.error.messagebox.showerror')
    def test_campos_excel_e_pdf_estao_vazios(self, mock_showerror):
        self.campo_arquivo_excel.text.return_value = ''
        self.campo_arquivo_pdf.text.return_value = ''
        result_existe_erro = existe_erro(self.campo_arquivo_excel.text(), self.campo_arquivo_pdf.text())
        mock_showerror.assert_called_once()
        self.assertIn('Por favor, selecione o arquivo Excel e o arquivo PDF.', mock_showerror.call_args.args[1])
        self.assertTrue(result_existe_erro)

    @patch('realce.infra.error.messagebox.showerror')
    def test_extensao_do_arquivo_nao_eh_suportada(self, mock_showerror):
        self.campo_arquivo_excel.text.return_value = 'extensao_de_planilha_nao_suportada.docx'
        self.campo_arquivo_pdf.text.return_value = 'extensao_de_pdf_nao_suportada.docx'
        result_arquivo_invalido = is_valid_file(self.campo_arquivo_excel.text(), self.campo_arquivo_pdf.text())
        mock_showerror.assert_called_once()
        self.assertIn('Ocorreu um problema', mock_showerror.call_args.args[0])
        self.assertFalse(result_arquivo_invalido)

    @patch('realce.infra.error.messagebox.showerror')
    def test_extensao_do_arquivo_eh_valida(self, mock_showerror):
        self.campo_arquivo_excel.text.return_value = 'extensao_de_planilha_valida.xlsx'
        self.campo_arquivo_pdf.text.return_value = 'extensao_pdf_valida.pdf'
        result_arquivo_valido = is_valid_file(self.campo_arquivo_excel.text(), self.campo_arquivo_pdf.text())
        mock_showerror.assert_not_called()
        self.assertTrue(result_arquivo_valido)

    @patch('realce.infra.error.messagebox.showerror')
    def test_campos_nao_sao_validos_quando_o_caminho_do_arquivo_nao_eh_especificado(self, mock_showerror):
        self.campo_arquivo_excel.text.return_value = 'campo_invalido.xlsx'
        self.campo_arquivo_pdf.text.return_value = 'campo_invalido.pdf'
        result_campos_sao_validos = campos_sao_validos(self.campo_arquivo_excel.text(), self.campo_arquivo_pdf.text())
        mock_showerror.assert_called_once()
        self.assertFalse(result_campos_sao_validos)

    @patch('realce.infra.error.os.path.isfile')
    @patch('realce.infra.error.messagebox.showerror')
    def test_quando_os_campos_sao_validos(self, mock_showerror, mock_isfile):
        mock_isfile.return_value = True
        self.campo_arquivo_excel.text.return_value = 'campo_valido.xlsx'
        self.campo_arquivo_pdf.text.return_value = 'campo_valido.pdf'
        result_campos_sao_validos = campos_sao_validos(self.campo_arquivo_excel.text(), self.campo_arquivo_pdf.text())
        mock_showerror.assert_not_called()
        self.assertTrue(result_campos_sao_validos)
