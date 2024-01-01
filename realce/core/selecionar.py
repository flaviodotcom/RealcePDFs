from tkinter import filedialog


def selecionar_arquivo_excel(campo_arquivo_excel):
    caminho_arquivo = filedialog.askopenfilename(filetypes=[("Arquivos Excel", "*.xlsx")])
    campo_arquivo_excel.clear()
    campo_arquivo_excel.insert(caminho_arquivo)


def selecionar_arquivo_pdf(campo_arquivo_pdf):
    caminho_arquivo = filedialog.askopenfilename(filetypes=[("Arquivos PDF", "*.pdf")])
    campo_arquivo_pdf.clear()
    campo_arquivo_pdf.insert(caminho_arquivo)
