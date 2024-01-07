import logging
import os
import sys
import webbrowser
from datetime import datetime

from PySide6.QtCore import Slot, Qt, QSize, QPoint
from PySide6.QtGui import QAction, QActionGroup, QIcon
from PySide6.QtWidgets import QMainWindow, QApplication, QWidget, QVBoxLayout, QMenu, QHBoxLayout, QFormLayout, \
    QGroupBox, QPushButton, QLineEdit, QStatusBar, QProgressBar, QToolTip
from qdarktheme import setup_theme

from realce.app.info import Tutorial
from realce.core.salvar import salvar_para_pasta_padrao, salvar_para_pasta_selecionada, salvar_para_pasta_do_vt
from realce.core.selecionar import selecionar_arquivo_excel, selecionar_arquivo_pdf
from realce.infra.helper import resource_path, atalhos
from realce.infra.logger import RealceLogger


class GuiHandler(logging.Handler):
    def __init__(self, output: QStatusBar):
        super().__init__()
        self.output = output

    def emit(self, record):
        self.output.showMessage(record.getMessage(), 10000)


class MainHome(QMainWindow):
    salvar: QPushButton
    salvar_como: QPushButton
    separar_vts_button: QPushButton
    cancelar_button: QPushButton
    progress_bar: QProgressBar
    excel_file: QLineEdit
    pdf_file: QLineEdit
    tutorial = None

    def __init__(self):
        super().__init__()
        self.setWindowTitle('RealcePDFs')
        self.setWindowIcon(QIcon(resource_path('resources/images/Cookie-Monster.ico')))
        self.setFixedSize(QSize(620, 280))

        self.logger = RealceLogger.get_logger()
        self.console_output = RealceLogger.console_output
        self.build_menu_bar()
        self.build_layout()
        atalhos(self)

    def init_theme_menu(self):
        menu = QMenu('Mudar tema')
        setup_theme('auto')
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
            setup_theme(custom_colors={"primary": "#D0BCFF"}, additional_qss=MainHome.qss_tooltip())

    def build_menu_bar(self):
        menu_bar = self.menuBar()

        menu_como_funciona = menu_bar.addMenu('&Tutorial')
        q_action = QAction(QIcon(), 'Funcionamento', self)
        q_action.triggered.connect(self.abrir_info)
        menu_como_funciona.addAction(q_action)

        menu_theme = menu_bar.addMenu('&Tema')
        menu_theme.addMenu(self.init_theme_menu())

        menu_log = menu_bar.addMenu('&Log')
        q_action = QAction(QIcon(), 'Gerar Log', self)
        q_action.triggered.connect(self.txt_log)
        menu_log.addAction(q_action)

    def build_layout(self):
        widget = QWidget(self)
        self.setCentralWidget(widget)

        output_layout = QVBoxLayout()
        output_layout.addWidget(self.build_log_output())
        output_layout.addWidget(self.build_progress_bar())
        output_layout.setSpacing(0)

        realce_layout = QVBoxLayout()
        realce_layout.addWidget(self.build_form())
        realce_layout.addWidget(self.build_separar_vts())
        realce_layout.addLayout(output_layout)
        realce_layout.setSpacing(8)

        main_layout = QHBoxLayout()
        main_layout.addLayout(realce_layout)
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
        self.salvar, self.salvar_como, encontrar_excel, encontrar_pdf = self.load_buttons()

        self.excel_file.setPlaceholderText('Insira o arquivo Excel (.xlsx, .xlsm, .xltx, .xltm)')
        self.pdf_file.setPlaceholderText('Insira o arquivo PDF (.pdf)')

        excel_layout, pdf_layout = QHBoxLayout(), QHBoxLayout()
        excel_layout.addWidget(self.excel_file)
        excel_layout.addWidget(encontrar_excel)
        acoes_form.addRow('Arquivo Excel:', excel_layout)

        pdf_layout.addWidget(self.pdf_file)
        pdf_layout.addWidget(encontrar_pdf)
        acoes_form.addRow('Arquivo PDF:', pdf_layout)

        layout = QHBoxLayout()
        layout.addWidget(self.salvar)
        layout.addWidget(self.salvar_como)
        acoes_form.addRow(layout)

        form_group = QGroupBox()
        form_group.setLayout(acoes_form)
        form_group.setFocusPolicy(Qt.FocusPolicy.ClickFocus)

        return form_group

    def build_separar_vts(self):
        separar_vts_form = QFormLayout()
        self.separar_vts_button = QPushButton("Separar PDFs")
        self.separar_vts_button.clicked.connect(lambda: salvar_para_pasta_do_vt(*self.get_field_text(), parent=self))

        self.cancelar_button = QPushButton('X')
        self.cancelar_button.setStyleSheet('background-color: #cc0000; color: white; font-weight: bold;')

        layout_vt = QHBoxLayout()
        layout_vt.addWidget(self.separar_vts_button, 1)
        layout_vt.addWidget(self.cancelar_button, alignment=Qt.AlignmentFlag.AlignRight)
        separar_vts_form.addRow(layout_vt)
        separar_vts_form.setSpacing(2)
        form_group_vt = QGroupBox("Vale Transporte")
        form_group_vt.setLayout(separar_vts_form)

        return form_group_vt

    def build_log_output(self):
        bar = QStatusBar()
        bar.showMessage("Olá!", 5000)
        bar.setStyleSheet("font-weight: bold; border: 1px solid;")
        self.logger.addHandler(GuiHandler(bar))
        return bar

    def build_progress_bar(self):
        self.progress_bar = QProgressBar()
        self.progress_bar.setStyleSheet("""
            QProgressBar {
                border: 1px solid;
                border-radius: 0px;
                border-top: none;
            }

            QProgressBar::chunk {
                background-color: #008000;
                border-radius: 0px;
            }

            .custom-color::chunk {
                background-color: #cc0000;
            }
        """)
        self.progress_bar.setTextVisible(False)
        self.progress_bar.setFixedHeight(10)
        return self.progress_bar

    def load_buttons(self):
        salvar, salvar_como = QPushButton("Salvar"), QPushButton("Salvar Como")
        salvar.clicked.connect(lambda: salvar_para_pasta_padrao(*self.get_field_text(), parent=self))
        salvar_como.clicked.connect(lambda: salvar_para_pasta_selecionada(*self.get_field_text(), parent=self))

        encontrar_excel, encontrar_pdf = QPushButton("Escolher arquivo"), QPushButton("Escolher arquivo")
        encontrar_excel.clicked.connect(lambda: selecionar_arquivo_excel(self.excel_file))
        encontrar_pdf.clicked.connect(lambda: selecionar_arquivo_pdf(self.pdf_file))

        return salvar, salvar_como, encontrar_excel, encontrar_pdf

    def msg_tooltip(self, text):
        tooltip = QToolTip()
        tooltip.showText(self.mapToGlobal(QPoint(200, 150)), text)

    @staticmethod
    def qss_tooltip():
        return """
        QToolTip {
                border-width: 1px;
                border-style: hidden;
                background-color: #292929;
            }
            """

    def abrir_info(self):
        if not (self.tutorial and self.tutorial.is_visible()):
            self.tutorial = Tutorial()
            self.tutorial.show()

    def txt_log(self):
        log_dir = resource_path('resources/logs')

        if not os.path.exists(log_dir):
            os.makedirs(log_dir)

        date = datetime.today().strftime(f"%d-%m-%Y_%Hh%M")
        log = resource_path(os.path.join(log_dir, f'log-{date}.txt'))

        with open(log, 'w') as log_file:
            log_file.write(self.console_output.getvalue())
            self.logger.info(f"Arquivo salvo em {os.path.realpath(log_file.name)}")
            webbrowser.open(log_dir)

    def get_field_text(self):
        return self.excel_file.text(), self.pdf_file.text()


class RunHome:
    def __init__(self):
        app = QApplication(sys.argv)
        main_home = MainHome()
        main_home.show()
        sys.exit(app.exec())
