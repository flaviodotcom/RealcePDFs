import sys
import logging

from PySide6.QtCore import Slot, Qt, QSize, QThread, Signal
from PySide6.QtGui import QAction, QActionGroup, QIcon
from PySide6.QtWidgets import QMainWindow, QApplication, QWidget, QVBoxLayout, QMenu, QHBoxLayout, QFormLayout, \
    QGroupBox, QPushButton, QLineEdit, QStatusBar, QProgressBar
from qdarktheme import setup_theme

from realce import RealceLogger
from realce.app import atalhos
from realce.app.info import Tutorial
from realce.core.salvar import salvar_para_pasta_padrao, salvar_para_pasta_selecionada
from realce.core.selecionar import SelectFiles
from realce.core.vt import SepararPDF
from realce.infra.helper import resource_path


class WorkerThread(QThread):
    progressUpdated = Signal(int)
    finished = Signal(object)

    def __init__(self, function_to_run, *args, priority=QThread.Priority.HighestPriority, **kwargs):
        super().__init__()
        self.function_to_run = function_to_run
        self.args = args
        self.kwargs = kwargs
        self.priority = priority

    def run(self):
        try:
            self.currentThread().setPriority(self.priority)
            result = self.function_to_run(*self.args, **self.kwargs)
            self.finished.emit(result)
        except Exception as e:
            RealceLogger.get_logger().error(e)
            if str(e).startswith('The SignalInstance object was already deleted'):
                RealceLogger.get_logger().info('Ocorreu um erro :C. Por favor, tente novamente.')
            self.finished.emit(e)


class GuiHandler(logging.Handler):
    def __init__(self, output: QStatusBar):
        super().__init__()
        self.output = output

    def emit(self, record):
        self.output.showMessage(record.getMessage(), 10000)


class MainHome(QMainWindow):
    separar_vts_button: QPushButton
    salvar: QPushButton
    salvar_como: QPushButton
    progress_bar: QProgressBar
    excel_file: QLineEdit
    pdf_file: QLineEdit
    tutorial = None
    thread_salvar_padrao = None
    thread_salvar_selecionada = None
    thread_vt = None

    def __init__(self):
        super().__init__()
        self.setWindowTitle('RealcePDFs')
        self.setWindowIcon(QIcon(resource_path('resources/images/Cookie-Monster.ico')))
        self.setFixedSize(QSize(620, 280))

        self.logger = RealceLogger.get_logger()
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
            setup_theme(custom_colors={"primary": "#D0BCFF"}) if theme == 'dark' else None

    def build_menu_bar(self):
        menu_bar = self.menuBar()

        menu_como_funciona = menu_bar.addMenu('&Tutorial')
        q_action = QAction(QIcon(), 'Funcionamento', self)
        q_action.triggered.connect(self.abrir_info)
        menu_como_funciona.addAction(q_action)

        menu_theme = menu_bar.addMenu('&Tema')
        menu_theme.addMenu(self.init_theme_menu())

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

        return form_group

    def build_separar_vts(self):
        separar_vts_form = QFormLayout()
        self.separar_vts_button = QPushButton("Separar PDFs")
        self.separar_vts_button.clicked.connect(lambda: self.acao_vt())

        layout_vt = QHBoxLayout()
        layout_vt.addWidget(self.separar_vts_button)
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
        """)
        self.progress_bar.setTextVisible(False)
        self.progress_bar.setFixedHeight(10)
        return self.progress_bar

    def load_buttons(self):
        salvar, salvar_como = QPushButton("Salvar"), QPushButton("Salvar Como")
        salvar.clicked.connect(lambda: self.acao_salvar_padrao())
        salvar_como.clicked.connect(lambda: self.acao_salvar_selecionada())

        encontrar_excel, encontrar_pdf = QPushButton("Escolher arquivo"), QPushButton("Escolher arquivo")
        encontrar_excel.clicked.connect(lambda: SelectFiles.selecionar_arquivo_excel(self.excel_file))
        encontrar_pdf.clicked.connect(lambda: SelectFiles.selecionar_arquivo_pdf(self.pdf_file))

        return salvar, salvar_como, encontrar_excel, encontrar_pdf

    def abrir_info(self):
        if not (self.tutorial and self.tutorial.is_visible()):
            self.tutorial = Tutorial()
            self.tutorial.show()

    def acao_salvar_padrao(self):
        self.thread_salvar_padrao = WorkerThread(salvar_para_pasta_padrao, self.excel_file, self.pdf_file)
        self.salvar.setEnabled(False)
        self.thread_salvar_padrao.finished.connect(lambda: self.handle_thread_finished(self.salvar))
        self.thread_salvar_padrao.progressUpdated.connect(self.update_progress_bar)
        self.thread_salvar_padrao.start()

    def acao_salvar_selecionada(self):
        self.thread_salvar_selecionada = WorkerThread(salvar_para_pasta_selecionada, self.excel_file, self.pdf_file)
        self.salvar_como.setEnabled(False)
        self.thread_salvar_selecionada.finished.connect(lambda: self.handle_thread_finished(self.salvar_como))
        self.thread_salvar_selecionada.progressUpdated.connect(self.update_progress_bar)
        self.thread_salvar_selecionada.start()

    def acao_vt(self):
        self.thread_vt = WorkerThread(SepararPDF.separar_vt, self.excel_file, self.pdf_file)
        self.separar_vts_button.setEnabled(False)
        self.thread_vt.finished.connect(lambda: self.handle_thread_finished(self.separar_vts_button))
        self.thread_vt.progressUpdated.connect(self.update_progress_bar)
        self.thread_vt.start()

    def update_progress_bar(self, value):
        self.progress_bar.setValue(value)

    @staticmethod
    def handle_thread_finished(button):
        button.setEnabled(True)


class RunHome:
    def __init__(self):
        app = QApplication(sys.argv)
        main_home = MainHome()
        main_home.show()
        sys.exit(app.exec())
