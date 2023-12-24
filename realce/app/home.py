import sys
import logging
import platform

from PySide6.QtCore import Slot, Qt, QSize
from PySide6.QtGui import QAction, QActionGroup, QIcon
from PySide6.QtWidgets import QMainWindow, QApplication, QWidget, QVBoxLayout, QMenu, QHBoxLayout, QFormLayout, \
    QGroupBox, QPushButton, QLineEdit, QStatusBar
from qdarktheme import setup_theme

from realce import RealceLogger
from realce.app import atalhos
from realce.app.info import Tutorial
from realce.core.salvar import salvar_para_pasta_padrao, salvar_para_pasta_selecionada
from realce.core.selecionar import SelectFiles
from realce.core.vt import SepararPDF
from realce.infra.helper import resource_path


class GuiHandler(logging.Handler):
    def __init__(self, output: QStatusBar):
        super().__init__()
        self.output = output

    def emit(self, record):
        self.output.showMessage(record.getMessage(), 10000)


class MainHome(QMainWindow):
    excel_file: QLineEdit
    pdf_file: QLineEdit
    tutorial = None

    def __init__(self):
        super().__init__()
        self.theme_is_w10 = self.init_theme_menu()
        self.setWindowTitle('RealcePDFs')
        self.setWindowIcon(QIcon(resource_path('resources/images/Cookie-Monster.ico')))
        self.setFixedSize(QSize(620, 300))

        self.logger = RealceLogger.get_logger()
        self.build_menu_bar()
        self.build_layout()
        atalhos(self)

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
        vertical_layout.addWidget(self.build_separar_vts())
        vertical_layout.addWidget(self.build_log_output())
        vertical_layout.setSpacing(25)

        main_layout = QHBoxLayout()
        main_layout.addLayout(vertical_layout)
        base_layout = QVBoxLayout(widget)
        base_layout.addLayout(main_layout)
        base_layout.setSpacing(15)

    def build_form(self):
        acoes_form = QFormLayout()
        acoes_form.setRowWrapPolicy(QFormLayout.DontWrapRows)
        acoes_form.setFieldGrowthPolicy(QFormLayout.ExpandingFieldsGrow)
        acoes_form.setLabelAlignment(Qt.AlignmentFlag.AlignLeft)
        acoes_form.setSpacing(8)

        self.excel_file, self.pdf_file = QLineEdit(), QLineEdit()
        salvar, salvar_como, encontrar_excel, encontrar_pdf = self.load_buttons()

        excel_layout, pdf_layout = QHBoxLayout(), QHBoxLayout()
        excel_layout.addWidget(self.excel_file)
        excel_layout.addWidget(encontrar_excel)
        acoes_form.addRow('Arquivo Excel:', excel_layout)

        pdf_layout.addWidget(self.pdf_file)
        pdf_layout.addWidget(encontrar_pdf)
        acoes_form.addRow('Arquivo PDF:', pdf_layout)

        layout = QHBoxLayout()
        layout.addWidget(salvar)
        layout.addWidget(salvar_como)
        acoes_form.addRow(layout)

        form_group = QGroupBox()
        form_group.setLayout(acoes_form)

        return form_group

    def build_separar_vts(self):
        separar_vts_form = QFormLayout()
        separar_vts_button = QPushButton("Separar PDFs")
        separar_vts_button.clicked.connect(lambda: SepararPDF.separar_vt(self.excel_file, self.pdf_file))

        layout_vt = QHBoxLayout()
        layout_vt.addWidget(separar_vts_button)
        separar_vts_form.addRow(layout_vt)
        form_group_vt = QGroupBox("Vale Transporte")
        form_group_vt.setLayout(separar_vts_form)

        return form_group_vt

    def build_log_output(self):
        bar = QStatusBar()
        bar.showMessage("Olá!", 5000)
        bar.setStyleSheet("font-weight: bold; border: 1px solid;")
        self.logger.addHandler(GuiHandler(bar))
        return bar

    def load_buttons(self):
        salvar, salvar_como = QPushButton("Salvar"), QPushButton("Salvar Como")
        salvar.clicked.connect(lambda: salvar_para_pasta_padrao(self.excel_file, self.pdf_file))
        salvar_como.clicked.connect(lambda: salvar_para_pasta_selecionada(self.excel_file, self.pdf_file))

        encontrar_excel, encontrar_pdf = QPushButton("Escolher arquivo"), QPushButton("Escolher arquivo")
        encontrar_excel.clicked.connect(lambda: SelectFiles.selecionar_arquivo_excel(self.excel_file))
        encontrar_pdf.clicked.connect(lambda: SelectFiles.selecionar_arquivo_pdf(self.pdf_file))

        return salvar, salvar_como, encontrar_excel, encontrar_pdf

    def abrir_info(self):
        if not (self.tutorial and self.tutorial.is_visible()):
            self.tutorial = Tutorial()
            self.tutorial.show()


class RunHome:
    def __init__(self):
        app = QApplication(sys.argv)
        main_home = MainHome()
        main_home.show()
        sys.exit(app.exec())
