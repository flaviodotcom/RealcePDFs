import os

from tkinter import messagebox, filedialog

from realce.core.destacar import RealceMatriculas
from realce.infra.error import tratar_erro, tratar_erro_excel, tratar_erro_pdf, ErroExcel, ErroPdf


def salvar_para_pasta_padrao(campo_arquivo_excel, campo_arquivo_pdf):
    caminho_arquivo_excel = campo_arquivo_excel.get()
    caminho_arquivo_pdf = campo_arquivo_pdf.get()

    try:
        tratar_erro(caminho_arquivo_excel, caminho_arquivo_pdf)
        pasta_destino = os.path.join(
            os.path.expanduser("~"), "Desktop", "BENEFICIOS DESTACADOS"
        )
        tratar_erro_pdf(caminho_arquivo_pdf)
        tratar_erro_excel(caminho_arquivo_excel)

        os.makedirs(pasta_destino, exist_ok=True)
        RealceMatriculas.pdf(pasta_destino, campo_arquivo_excel, campo_arquivo_pdf)
    except (ErroExcel, ErroPdf) as e:
        messagebox.showerror("Erro", str(e))


def salvar_para_pasta_selecionada_pelo_usuario(campo_arquivo_excel, campo_arquivo_pdf):
    caminho_arquivo_excel = campo_arquivo_excel.get()
    caminho_arquivo_pdf = campo_arquivo_pdf.get()

    if tratar_erro(caminho_arquivo_excel, caminho_arquivo_pdf):
        return

    pasta_destino = filedialog.askdirectory()
    if not pasta_destino:
        return
    RealceMatriculas.pdf(pasta_destino, campo_arquivo_excel, campo_arquivo_pdf)
