import sys

from PySide6.QtCore import Qt
from PySide6.QtGui import QIcon
from PySide6.QtWidgets import QMainWindow, QApplication, QWidget, QVBoxLayout, QPushButton, QLabel

from realce.infra.helper import resource_path, atalhos


class Tutorial(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Como Funciona?")
        self.setFixedSize(620, 500)
        self.setWindowIcon(QIcon(resource_path('resources/images/Taz.ico')))
        self.setWindowFlag(Qt.WindowType.WindowCloseButtonHint, True)

        self.tutorial()
        atalhos(self)

    def tutorial(self):
        central_widget = QWidget(self)
        self.setCentralWidget(central_widget)

        layout = QVBoxLayout(central_widget)

        texto_informacoes = (
            "Este programa permite destacar a matrícula dos funcionários em um arquivo PDF usando informações de "
            "uma planilha do Excel."
            " Principalmente usado para realçar os benefícios de Seguro de Vida, Plano Odontológico, "
            "Vale Transporte, Vale Alimentação e Vale Refeição.\n\n"
            "Orientações:\n\n"
            "1. Clique no botão 'Selecionar' para escolher o arquivo Excel e PDF, respectivamente.\n\n"
            "2. Certifique-se de que as matrículas estejam na coluna B da planilha. O programa percorre pela "
            "segunda coluna (coluna B),"
            " garanta que nessa coluna não existam outras informações.\n\n"
            "3. O programa cria um arquivo de texto, que aponta quais foram as matrículas não encontradas junto "
            "com o nome do colaborador."
            " Para que essa funcionalidade ocorra como esperado, mantenha o nome dos colaboradores na terceira "
            "coluna (coluna C) da planilha."
            " O arquivo de texto será salvo no mesmo diretório do PDF editado.\n\n"
            "4. Após selecionar os arquivos, clique no botão 'Salvar' para guardar o PDF editado na Área de "
            "Trabalho (Desktop),"
            "dentro da pasta 'BENEFÍCIOS DESTACADOS'. Alternativamente, clique no botão 'Salvar Como' para "
            "escolher o local de armazenamento"
            "do PDF editado que preferir.\n\n"
            "Separar PDFS:\n\n"
            "1. Clique no botão 'Selecionar' para escolher o arquivo Excel e PDF, respectivamente.\n\n"
            "2. Após selecionar os arquivos, clique no botão 'Separar PDFs por Matricula' para escolher a pasta "
            "de destino do PDF realçado,"
            "logo após o programa irá criar uma pasta chamada 'Vt Separados', onde irá conter cada pagina do PDF "
            "realçado que encontrou a"
            "matrícula, e no fim irá juntar todos os PDFs que estão na pasta em um arquivo chamado 'Vt Final'."
        )

        rotulo_informacoes = QLabel(texto_informacoes)
        rotulo_informacoes.setWordWrap(True)
        rotulo_informacoes.setAlignment(Qt.AlignmentFlag.AlignTop | Qt.AlignmentFlag.AlignLeft)
        layout.addWidget(rotulo_informacoes, alignment=Qt.AlignmentFlag.AlignTop)

        fechar_button = QPushButton("Fechar", self)
        fechar_button.clicked.connect(self.fechar_janela)
        layout.addWidget(fechar_button, alignment=Qt.AlignmentFlag.AlignBottom | Qt.AlignmentFlag.AlignRight)

    def fechar_janela(self):
        self.close()

    def is_visible(self):
        return self.isVisible()


if __name__ == "__main__":
    app = QApplication(sys.argv)
    main_home = Tutorial()
    main_home.show()
    sys.exit(app.exec())
