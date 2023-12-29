import os

from tkinter import filedialog

from realce.core.destacar import RealceMatriculas
from realce.infra.error import campos_sao_validos


def salvar_para_pasta_padrao(campo_arquivo_excel, campo_arquivo_pdf):
    caminho_arquivo_excel = campo_arquivo_excel.text()
    caminho_arquivo_pdf = campo_arquivo_pdf.text()

    if campos_sao_validos(caminho_arquivo_excel, caminho_arquivo_pdf):
        pasta_destino = os.path.join(os.path.expanduser("~"), "Desktop", "BENEFICIOS DESTACADOS")
        os.makedirs(pasta_destino, exist_ok=True)
        RealceMatriculas.pdf(pasta_destino, campo_arquivo_excel, campo_arquivo_pdf)


def salvar_para_pasta_selecionada(campo_arquivo_excel, campo_arquivo_pdf):
    caminho_arquivo_excel = campo_arquivo_excel.text()
    caminho_arquivo_pdf = campo_arquivo_pdf.text()

    if campos_sao_validos(caminho_arquivo_excel, caminho_arquivo_pdf):
        pasta_destino = filedialog.askdirectory()
        if pasta_destino:
            RealceMatriculas.pdf(pasta_destino, campo_arquivo_excel, campo_arquivo_pdf)
