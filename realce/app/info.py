import sys

from PySide6.QtCore import Qt
from PySide6.QtGui import QIcon
from PySide6.QtWidgets import QMainWindow, QApplication, QWidget, QVBoxLayout, QPushButton, QLabel

from realce.infra.helper import resource_path, atalhos


class Tutorial(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Como Funciona?")
        self.setFixedSize(620, 520)
        self.setWindowIcon(QIcon(resource_path('resources/images/Taz.ico')))
        self.setWindowFlag(Qt.WindowType.WindowCloseButtonHint, True)

        self.tutorial()
        atalhos(self)

    def tutorial(self):
        central_widget = QWidget(self)
        self.setCentralWidget(central_widget)

        layout = QVBoxLayout(central_widget)

        info_text = (
            """
<h1 align="center">RealcePDFs</h1>
            
Este programa permite destacar a matrícula dos funcionários em um arquivo PDF usando informações de uma
planilha do Excel. Principalmente usado para destacar os benefícios de Seguro de Vida, Plano Odontológico,
Vale Transporte, Vale Alimentação e Vale Refeição.

<h2>Orientações:</h2>

<h3>Destaque</h3>

1. Clique no botão 'Selecionar' para escolher o arquivo Excel e PDF, respectivamente.<br>
2. Certifique-se de que as matrículas estejam na coluna B da planilha. O programa percorre pela segunda
coluna (coluna B), garanta que nessa coluna não existam outras informações.<br>
3. O programa cria um arquivo de texto, que aponta quais foram as matrículas não encontradas junto com o
nome do colaborador.<br>
4. Após selecionar os arquivos, clique no botão 'Salvar' para guardar o PDF editado na Área de Trabalho
(Desktop), dentro da pasta 'BENEFÍCIOS DESTACADOS'. Alternativamente, clique no botão 'Salvar Como' para 
escolher o local de armazenamento do PDF editado que preferir.

<h3>Separar PDFs</h3>

1. Clique no botão 'Selecionar' para escolher o arquivo Excel e PDF, respectivamente.<br>
2. Após selecionar os arquivos, clique no botão 'Separar PDFs por Matricula' para escolher a pasta de
destino do PDF realçado, após selecionar o caminho de destino, o programa irá criar uma pasta chamada 'Vt',
onde irá conter cada pagina do PDF destacado em que uma matrícula foi encontrada. Finalmente, todos os PDFs
destacados serão juntados em um arquivo único, sem duplicações (considerar mais de uma matrícula por PDF),
este arquivo é nomeado de 'PDF_MERGEADO'.
            """
        )

        label_tutorial = QLabel()
        label_tutorial.setMargin(10)
        label_tutorial.setText(info_text)
        label_tutorial.setWordWrap(True)
        label_tutorial.setTextFormat(Qt.TextFormat.RichText)
        label_tutorial.setTextInteractionFlags(Qt.TextInteractionFlag.TextSelectableByMouse)
        layout.addWidget(label_tutorial)

        fechar_button = QPushButton("Fechar", self)
        fechar_button.clicked.connect(self.fechar_janela)
        layout.addWidget(fechar_button, alignment=Qt.AlignmentFlag.AlignBottom | Qt.AlignmentFlag.AlignCenter)

    def fechar_janela(self):
        self.close()

    def is_visible(self):
        return self.isVisible()


if __name__ == "__main__":
    app = QApplication(sys.argv)
    main_home = Tutorial()
    main_home.show()
    sys.exit(app.exec())
