import os
from tkinter import filedialog

from realce.core.destacar import RealceMatriculas
from realce.core.vt import SepararPDF
from realce.infra.error import campos_sao_validos
from realce.infra.thread import WorkerThread


def salvar_para_pasta_padrao(campo_arquivo_excel, campo_arquivo_pdf, parent=None):
    ThreadManager(parent, RealceMatriculas.pdf, campo_arquivo_excel, campo_arquivo_pdf, parent.salvar_como, True)


def salvar_para_pasta_selecionada(campo_arquivo_excel, campo_arquivo_pdf, parent=None):
    ThreadManager(parent, RealceMatriculas.pdf, campo_arquivo_excel, campo_arquivo_pdf, parent.salvar_como)


def salvar_para_pasta_do_vt(campo_arquivo_excel, campo_arquivo_pdf, parent=None):
    ThreadManager(parent, SepararPDF.separar_vt, campo_arquivo_excel, campo_arquivo_pdf, parent.separar_vts_button)


class ThreadManager:
    def __init__(self, parent, function, campo_arquivo_excel, campo_arquivo_pdf, button, default=None):
        self.parent = parent
        self.function = function
        self.campo_arquivo_excel = campo_arquivo_excel
        self.campo_arquivo_pdf = campo_arquivo_pdf
        self.button = button
        self.default = default
        self.configurar_thread()

    def configurar_thread(self):
        if campos_sao_validos(self.campo_arquivo_excel, self.campo_arquivo_pdf):
            pasta_destino = self.obter_pasta_destino()
            if pasta_destino:
                thread = WorkerThread(self.function, self.campo_arquivo_excel, self.campo_arquivo_pdf, pasta_destino,
                                      parent=self.parent)
                self.configurar_botao(thread)
                thread.start()

    def configurar_botao(self, thread):
        self.button.setEnabled(False)
        self.parent.cancelar_button.clicked.connect(lambda: thread.stop_execution(self.button))
        thread.finished.connect(lambda: thread.handle_thread_finished(self.button))
        thread.progressUpdated.connect(thread.update_progress_bar)

    def obter_pasta_destino(self):
        if self.default:
            pasta_destino = os.path.join(os.path.expanduser("~"), "Desktop", "BENEFICIOS DESTACADOS")
            os.makedirs(pasta_destino, exist_ok=True)
            return pasta_destino
        return filedialog.askdirectory()
