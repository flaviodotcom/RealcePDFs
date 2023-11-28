import os

import fitz
import openpyxl
from pathlib import Path
from tkinter import messagebox

from realce.core.selecionar import SelectFiles


class BaseRealcePdf:

    @staticmethod
    def destacar_pdf(campo_arquivo_excel, campo_arquivo_pdf):
        nome_arquivo = SelectFiles.guardar_nome()
        caminho_arquivo_excel = campo_arquivo_excel.get()
        caminho_arquivo_pdf = campo_arquivo_pdf.get()

        arquivo_excel = openpyxl.load_workbook(caminho_arquivo_excel)
        arquivo_pdf = fitz.open(caminho_arquivo_pdf)

        matriculas_nao_encontradas = []

        for linha in range(1, arquivo_excel.active.max_row + 1):
            numero_matricula = str(arquivo_excel.active.cell(row=linha, column=2).value)
            nome_matricula = str(arquivo_excel.active.cell(row=linha, column=3).value).upper()

            encontrou_matricula = False

            for pagina in arquivo_pdf:
                for linha_texto in pagina.get_text().splitlines():
                    if numero_matricula in linha_texto:
                        realce = pagina.search_for(numero_matricula, hit_max=1)
                        if realce:
                            retangulo_realce = fitz.Rect(realce[0][:4])
                            pagina.add_highlight_annot(retangulo_realce)
                            encontrou_matricula = True
                            break

            if not encontrou_matricula and nome_matricula and numero_matricula != "None":
                matriculas_nao_encontradas.append(f"{numero_matricula} - {nome_matricula}")

        return arquivo_pdf, matriculas_nao_encontradas, nome_arquivo


class RealceMatriculas(BaseRealcePdf):
    caminho_arquivo_saida: str | Path

    @staticmethod
    def pdf(pasta_destino, campo_arquivo_excel, campo_arquivo_pdf):
        arquivo_pdf, matriculas_nao_encontradas, nome_arquivo = BaseRealcePdf.destacar_pdf(campo_arquivo_excel,
                                                                                           campo_arquivo_pdf)

        RealceMatriculas.salvar_arquivo_pdf(pasta_destino, nome_arquivo, arquivo_pdf)
        RealceMatriculas.exibir_resultados(matriculas_nao_encontradas)

    @staticmethod
    def salvar_arquivo_pdf(pasta_destino, nome_arquivo, arquivo_pdf):
        nome_arquivo_saida = nome_arquivo
        numero_arquivo = 1

        while os.path.exists(os.path.join(pasta_destino, nome_arquivo_saida)):
            nome_arquivo_saida = f"{os.path.splitext(nome_arquivo)[0]}({numero_arquivo}).pdf"
            numero_arquivo += 1

        RealceMatriculas.caminho_arquivo_saida = os.path.join(pasta_destino, nome_arquivo_saida)
        arquivo_pdf.save(RealceMatriculas.caminho_arquivo_saida)

    @staticmethod
    def exibir_resultados(matriculas_nao_encontradas):
        messagebox.showinfo("Concluído", f"O PDF editado foi salvo em:\n{RealceMatriculas.caminho_arquivo_saida}")

        if matriculas_nao_encontradas:
            RealceMatriculas.salvar_matriculas_nao_encontradas(matriculas_nao_encontradas,
                                                               RealceMatriculas.caminho_arquivo_saida)

    @staticmethod
    def salvar_matriculas_nao_encontradas(matriculas_nao_encontradas, pasta_destino):
        nome_arquivo_txt = "Matrículas não encontradas.txt"
        numero_arquivo_txt = 1

        while os.path.exists(os.path.join(pasta_destino, nome_arquivo_txt)):
            nome_arquivo_txt = f"Matrículas não encontradas({numero_arquivo_txt}).txt"
            numero_arquivo_txt += 1

        caminho_arquivo_txt = os.path.join(os.path.dirname(pasta_destino), nome_arquivo_txt)

        with open(caminho_arquivo_txt, "w") as arquivo_txt:
            for matricula in matriculas_nao_encontradas:
                arquivo_txt.write(matricula + "\n")

        messagebox.showwarning(
            "Matrículas não encontradas",
            f"Não foi possível encontrar algumas matrículas no arquivo PDF selecionado.\n\nFoi gerado um "
            f"arquivo de texto que contém as matrículas não encontradas, salvo em:\n{caminho_arquivo_txt}",
        )
