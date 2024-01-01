import os
from tkinter import filedialog

from realce.core.destacar import RealceMatriculas
from realce.core.vt import SepararPDF
from realce.infra.error import campos_sao_validos
from realce.infra.thread import WorkerThread


def salvar_para_pasta_padrao(campo_arquivo_excel, campo_arquivo_pdf, parent=None):
    if campos_sao_validos(campo_arquivo_excel, campo_arquivo_pdf):
        pasta_destino = os.path.join(os.path.expanduser("~"), "Desktop", "BENEFICIOS DESTACADOS")
        os.makedirs(pasta_destino, exist_ok=True)

        thread_padrao = WorkerThread(RealceMatriculas.pdf, campo_arquivo_excel, campo_arquivo_pdf, pasta_destino,
                                     parent=parent)
        parent.salvar.setEnabled(False)
        thread_padrao.finished.connect(lambda: thread_padrao.handle_thread_finished(parent.salvar))
        thread_padrao.progressUpdated.connect(thread_padrao.update_progress_bar)
        thread_padrao.start()


def salvar_para_pasta_selecionada(campo_arquivo_excel, campo_arquivo_pdf, parent=None):
    if campos_sao_validos(campo_arquivo_excel, campo_arquivo_pdf):
        pasta_destino = filedialog.askdirectory()

        if pasta_destino:
            thread_selecionada = WorkerThread(RealceMatriculas.pdf, campo_arquivo_excel, campo_arquivo_pdf,
                                              pasta_destino, parent=parent)
            parent.salvar_como.setEnabled(False)
            thread_selecionada.finished.connect(lambda: thread_selecionada.handle_thread_finished(parent.salvar_como))
            thread_selecionada.progressUpdated.connect(thread_selecionada.update_progress_bar)
            thread_selecionada.start()


def salvar_para_pasta_do_vt(campo_arquivo_excel, campo_arquivo_pdf, parent=None):
    if campos_sao_validos(campo_arquivo_excel, campo_arquivo_pdf):
        pasta_destino = filedialog.askdirectory()

        if pasta_destino:
            thread_vt = WorkerThread(SepararPDF.separar_vt, campo_arquivo_excel, campo_arquivo_pdf, pasta_destino,
                                     parent=parent)
            parent.separar_vts_button.setEnabled(False)
            thread_vt.finished.connect(lambda: thread_vt.handle_thread_finished(parent.separar_vts_button))
            thread_vt.progressUpdated.connect(thread_vt.update_progress_bar)
            thread_vt.start()
