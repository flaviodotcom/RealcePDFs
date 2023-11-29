import os

from openpyxl import load_workbook
from PyPDF4 import PdfFileMerger
from PyPDF2 import PdfWriter, PdfReader
from tkinter import filedialog, messagebox

from realce.core.destacar import BaseRealcePdf, RealceMatriculas
from realce.infra.error import tratar_erro, tratar_pasta_destino, confirmar_diretorio


class SepararPDF(BaseRealcePdf):

    @staticmethod
    def separar_vt(campo_arquivo_excel, campo_arquivo_pdf):
        pasta_destino = SepararPDF.tratamento(campo_arquivo_excel.get(), campo_arquivo_pdf.get())
        arquivo_pdf, matriculas_nao_encontradas, nome_arquivo = SepararPDF.destacar_pdf(campo_arquivo_excel,
                                                                                        campo_arquivo_pdf)

        nome_arquivo = RealceMatriculas.salvar_arquivo_pdf(pasta_destino, nome_arquivo, arquivo_pdf)
        # nome_arquivo = SepararPDF.nome_arquivo_final(pasta_destino, nome_arquivo)
        # SepararPDF.salvar_arquivo_pdf(arquivo_pdf, pasta_destino, nome_arquivo)

        arquivo_pdf = PdfReader(os.path.join(pasta_destino, nome_arquivo))
        pasta_destino = SepararPDF.criar_pasta_vt_separado(pasta_destino)

        SepararPDF.separar_pdf_por_matricula(arquivo_pdf, campo_arquivo_excel, pasta_destino)

        pdf_mesclado = SepararPDF.mesclar_pdfs(pasta_destino)
        SepararPDF.salvar_arquivo_mesclado(pdf_mesclado, pasta_destino)

        SepararPDF.exibir_mensagem_conclusao(pasta_destino, matriculas_nao_encontradas)

    @staticmethod
    def get_num_linhas(campo_arquivo_excel):
        planilha = load_workbook(campo_arquivo_excel.get()).active
        return planilha.max_row

    # @staticmethod
    # def nome_arquivo_final(pasta_destino, nome_arquivo):
    #     numero_arquivo = 1
    #     while os.path.exists(os.path.join(pasta_destino, nome_arquivo)):
    #         nome_arquivo = f"{os.path.splitext(nome_arquivo)[0]}({numero_arquivo}).pdf"
    #         numero_arquivo += 1
    #     return nome_arquivo
    #
    # @staticmethod
    # def salvar_arquivo_pdf(arquivo_pdf, pasta_destino, nome_arquivo):
    #     caminho_arquivo_saida = os.path.join(pasta_destino, nome_arquivo)
    #     arquivo_pdf.save(caminho_arquivo_saida)

    @staticmethod
    def criar_pasta_vt_separado(pasta_destino):
        pasta_destino = f'{pasta_destino}/Vt Separado'
        check_folder = os.path.isdir(pasta_destino)
        if not check_folder:
            os.makedirs(pasta_destino)
        return pasta_destino

    @staticmethod
    def separar_pdf_por_matricula(arquivo_pdf, campo_arquivo_excel, pasta_destino):
        planilha = load_workbook(campo_arquivo_excel.get()).active
        for linha in range(1, SepararPDF.get_num_linhas(campo_arquivo_excel) + 1):
            numero_matricula = str(planilha.cell(row=linha, column=2).value)
            nome_func = str(planilha.cell(row=linha, column=3).value)
            new_pdf = PdfWriter()

            for pagina in arquivo_pdf.pages:
                if numero_matricula in pagina.extract_text():
                    new_pdf.add_page(pagina)

            if len(new_pdf.pages) > 0:
                output_file = f"{pasta_destino}/{nome_func}.pdf"
                with open(output_file, "wb") as f:
                    new_pdf.write(f)

    @staticmethod
    def mesclar_pdfs(pasta_destino):
        pdf_mesclado = PdfFileMerger()
        for arquivo in os.listdir(pasta_destino):
            if arquivo.lower().endswith(".pdf"):
                caminho_arquivo = os.path.join(pasta_destino, arquivo)
                with open(caminho_arquivo, 'rb') as f:
                    pdf_mesclado.append(f)
        return pdf_mesclado

    @staticmethod
    def salvar_arquivo_mesclado(pdf_mesclado, pasta_destino):
        if pdf_mesclado.pages:
            caminho_arquivo_mesclado = f"{pasta_destino}/- Vt final.pdf"
            with open(caminho_arquivo_mesclado, "wb") as saida:
                pdf_mesclado.write(saida)

    @staticmethod
    def tratamento(campo_arquivo_excel, campo_arquivo_pdf):
        if tratar_erro(campo_arquivo_excel, campo_arquivo_pdf):
            return False

        pasta_destino = filedialog.askdirectory()
        if not tratar_pasta_destino(pasta_destino):
            if confirmar_diretorio(pasta_destino):
                return pasta_destino
        return False

    @staticmethod
    def exibir_mensagem_conclusao(pasta_destino, matriculas_nao_encontradas):
        messagebox.showinfo("Concluído", f"O Arquivo final foi salvo em:\n{pasta_destino}")

        if matriculas_nao_encontradas:
            nome_arquivo_txt = "Matrículas não encontradas.txt"
            numero_arquivo_txt = 1

            while os.path.exists(os.path.join(pasta_destino, nome_arquivo_txt)):
                nome_arquivo_txt = f"Matrículas não encontradas({numero_arquivo_txt}).txt"
                numero_arquivo_txt += 1

            caminho_arquivo_txt = os.path.join(pasta_destino, nome_arquivo_txt)

            with open(caminho_arquivo_txt, "w") as arquivo_txt:
                for matricula in matriculas_nao_encontradas:
                    arquivo_txt.write(matricula + "\n")

            messagebox.showwarning(
                "Matrículas não encontradas",
                f"Não foi possível encontrar algumas matrículas no arquivo PDF selecionado.\n\nFoi gerado um "
                f"arquivo de texto que contém as matrículas não encontradas, salvo em:\n{caminho_arquivo_txt}",
            )
