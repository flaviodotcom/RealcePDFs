import os

from tkinter import filedialog


class SelectFiles:
    nome_arquivo: str

    def __init__(self, campo_arquivo_excel, campo_arquivo_pdf):
        self.selecionar_arquivo_excel(campo_arquivo_excel)
        self.selecionar_arquivo_pdf(campo_arquivo_pdf)

    @staticmethod
    def selecionar_arquivo_excel(campo_arquivo_excel):
        caminho_arquivo = filedialog.askopenfilename(filetypes=[("Arquivos Excel", "*.xlsx")])
        campo_arquivo_excel.clear()
        campo_arquivo_excel.insert(caminho_arquivo)

    @staticmethod
    def selecionar_arquivo_pdf(campo_arquivo_pdf):
        caminho_arquivo = filedialog.askopenfilename(filetypes=[("Arquivos PDF", "*.pdf")])
        campo_arquivo_pdf.clear()
        campo_arquivo_pdf.insert(caminho_arquivo)
        SelectFiles.nome_arquivo = os.path.basename(caminho_arquivo)

    @staticmethod
    def guardar_nome():
        return SelectFiles.nome_arquivo
