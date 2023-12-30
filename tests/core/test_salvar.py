import unittest
from unittest.mock import patch, MagicMock

from realce.core.salvar import salvar_para_pasta_padrao, salvar_para_pasta_selecionada


class SalvarTestCase(unittest.TestCase):
    def setUp(self):
        self.campo_arquivo_excel = MagicMock()
        self.campo_arquivo_pdf = MagicMock()
        self.campo_arquivo_excel.text.return_value = 'caminho/ou/arquivo/nao/existe/excel.xlsx'
        self.campo_arquivo_pdf.text.return_value = 'caminho/ou/arquivo/nao/existe/arquivo.pdf'

    def get_field_text(self):
        return self.campo_arquivo_excel.text(), self.campo_arquivo_pdf.text()

    @patch('realce.infra.error.messagebox.showerror')
    def test_nao_salvar_quando_path_nao_existe_padrao(self, mock_showerror):
        salvar_para_pasta_padrao(*self.get_field_text())
        mock_showerror.assert_called_once_with('Ocorreu um problema', 'Por favor, verifique o caminho do arquivo PDF.')

    @patch('realce.infra.error.messagebox.showerror')
    @patch('realce.core.salvar.filedialog.askdirectory')
    def test_nao_salvar_quando_path_nao_existe_selecionada(self, mock_askdirectory, mock_showerror):
        mock_askdirectory.return_value = 'C:/caminho/para/pasta_selecionada/inexistente'
        salvar_para_pasta_selecionada(*self.get_field_text())
        mock_showerror.assert_called_once_with('Ocorreu um problema', 'Por favor, verifique o caminho do arquivo PDF.')
