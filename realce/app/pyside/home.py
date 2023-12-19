import sys
import platform

from PySide6.QtCore import Slot, Qt, QSize
from PySide6.QtGui import QAction, QActionGroup, QIcon
from PySide6.QtWidgets import QMainWindow, QApplication, QWidget, QVBoxLayout, QMenu, QHBoxLayout, QFormLayout, \
    QGroupBox, QPushButton, QLineEdit

from qdarktheme import setup_theme

from realce.app.info import SegundaJanela
from realce.infra.helper import resource_path


class MainHome(QMainWindow):
    def __init__(self):
        super().__init__()
        self.theme_is_w10 = self.init_theme_menu()
        self.setWindowTitle('RealcePDFs')
        self.setWindowIcon(QIcon(resource_path('resources/images/Cookie-Monster.ico')))
        self.setMinimumSize(QSize(620, 400))

        self.build_layout()
        self.build_menu_bar()

    def init_theme_menu(self):
        so = dict(Linux=False, Darwin=False)
        show_combo_box = so.get(platform.system(), True)

        if show_combo_box and platform.release() == '10':
            menu = QMenu('Mudar tema')
            setup_theme("auto")
            theme_group = QActionGroup(self)
            themes = ['dark', 'light', 'auto']

            for theme in themes:
                action = QAction(theme.capitalize(), self, checkable=True)
                action.triggered.connect(lambda *args, get_theme=theme: self.toggle_theme(get_theme))
                theme_group.addAction(action)
                menu.addAction(action)

            return menu

    @staticmethod
    @Slot(str)
    def toggle_theme(theme) -> None:
        setup_theme(theme)
        if theme == 'dark':
            setup_theme(custom_colors={"primary": "#D0BCFF"}) if theme == 'dark' else None

    def build_menu_bar(self):
        menu_bar = self.menuBar()

        menu_como_funciona = menu_bar.addMenu('&Tutorial')
        q_action = QAction(QIcon(), 'Funcionamento', self)
        q_action.triggered.connect(self.abrir_info)
        menu_como_funciona.addAction(q_action)

        if self.theme_is_w10:
            menu_theme = menu_bar.addMenu('&Tema')
            menu_theme.addMenu(self.theme_is_w10)

    def build_layout(self):
        widget = QWidget(self)
        self.setCentralWidget(widget)

        vertical_layout = QVBoxLayout()
        vertical_layout.addWidget(self.build_form())
        vertical_layout.setSpacing(25)

        main_layout = QHBoxLayout()
        main_layout.addLayout(vertical_layout)
        base_layout = QVBoxLayout(widget)
        base_layout.addLayout(main_layout)
        base_layout.setSpacing(15)

    @staticmethod
    def build_form():
        acoes_form = QFormLayout()
        acoes_form.setRowWrapPolicy(QFormLayout.DontWrapRows)
        acoes_form.setFieldGrowthPolicy(QFormLayout.ExpandingFieldsGrow)
        acoes_form.setLabelAlignment(Qt.AlignmentFlag.AlignLeft)
        acoes_form.setSpacing(15)

        excel_file, pdf_file = QLineEdit(), QLineEdit()
        salvar, salvar_como = QPushButton("Salvar"), QPushButton("Salvar Como")

        acoes_form.addRow('Arquivo Excel', excel_file)
        acoes_form.addRow('Arquivo PDF', pdf_file)
        layout = QHBoxLayout()
        layout.addWidget(salvar)
        layout.addWidget(salvar_como)
        acoes_form.addRow(layout)

        form_group = QGroupBox()
        form_group.setLayout(acoes_form)

        return form_group

    @staticmethod
    def abrir_info():
        info = SegundaJanela()
        info.abrir_janela()


if __name__ == "__main__":
    app = QApplication(sys.argv)
    main_home = MainHome()
    main_home.show()
    sys.exit(app.exec())
