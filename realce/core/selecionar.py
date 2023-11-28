import os

from pathlib import Path
from tkinter import filedialog
from customtkinter import END


class SelectFiles:
    nome_arquivo: str | Path

    def __init__(self, campo_arquivo_excel, campo_arquivo_pdf):
        self.selecionar_arquivo_excel(campo_arquivo_excel)
        self.selecionar_arquivo_pdf(campo_arquivo_pdf)

    @staticmethod
    def selecionar_arquivo_excel(campo_arquivo_excel):
        caminho_arquivo = filedialog.askopenfilename(
            filetypes=[
                ("Arquivos Excel", "*.xlsx"),
                ("Arquivos CSV", "*.csv"),
                ("Arquivos Excel 97-2003", "*.xls"),
            ]
        )
        campo_arquivo_excel.delete(0, END)
        campo_arquivo_excel.insert(0, caminho_arquivo)

    @staticmethod
    def selecionar_arquivo_pdf(campo_arquivo_pdf):
        caminho_arquivo = filedialog.askopenfilename(filetypes=[("Arquivos PDF", "*.pdf")])
        campo_arquivo_pdf.delete(0, END)
        campo_arquivo_pdf.insert(0, caminho_arquivo)
        SelectFiles.nome_arquivo = os.path.basename(caminho_arquivo)

    @staticmethod
    def guardar_nome():
        return SelectFiles.nome_arquivo
