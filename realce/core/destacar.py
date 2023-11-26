import os

import fitz
import openpyxl
from tkinter import messagebox


class RealceMatriculas:

    def __init__(self, pasta_destino, campo_arquivo_excel, campo_arquivo_pdf):
        self.destacar_pdf(pasta_destino, campo_arquivo_excel, campo_arquivo_pdf)

    @staticmethod
    def destacar_pdf(pasta_destino, campo_arquivo_excel, campo_arquivo_pdf):
        nome_arquivo = ''

        caminho_arquivo_excel = campo_arquivo_excel.get()
        arquivo_excel = openpyxl.load_workbook(caminho_arquivo_excel)

        caminho_arquivo_pdf = campo_arquivo_pdf.get()
        arquivo_pdf = fitz.open(caminho_arquivo_pdf)

        planilha = arquivo_excel.active

        num_linhas = planilha.max_row

        matriculas_nao_encontradas = []

        for linha in range(1, num_linhas + 1):
            numero_matricula = planilha.cell(row=linha, column=2).value
            numero_matricula = str(numero_matricula)

            nome_matricula = planilha.cell(row=linha, column=3).value
            nome_matricula = str(nome_matricula).upper()

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
                matriculas_nao_encontradas.append(numero_matricula + " - " + nome_matricula)

        nome_arquivo_saida = nome_arquivo
        numero_arquivo = 1
        while os.path.exists(os.path.join(pasta_destino, nome_arquivo_saida)):
            nome_arquivo_saida = f"{nome_arquivo.strip('.pdf')}({numero_arquivo}).pdf"
            numero_arquivo += 1

        caminho_arquivo_saida = os.path.join(pasta_destino, nome_arquivo_saida)
        arquivo_pdf.save(caminho_arquivo_saida)

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

        messagebox.showinfo(
            "Concluído", f"O PDF editado foi salvo em:\n{caminho_arquivo_saida}"
        )

        caminho_arquivo_txt = None

        messagebox.showwarning(
            "Matrículas não encontradas",
            f"Não foi possível encontrar algumas matrículas no arquivo PDF selecionado.\n\nFoi gerado um "
            f"arquivo de texto que contém as matrículas não encontradas, salvo em:\n{caminho_arquivo_txt}",
        )
