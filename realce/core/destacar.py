import os
import re

import fitz
import openpyxl
from tkinter import messagebox

from realce import get_logger
from realce.core.selecionar import SelectFiles


class BaseRealcePdf:
    caminho_arquivo_saida: str
    logger = get_logger('RealcePDFs')

    @staticmethod
    def destacar_pdf(campo_arquivo_excel, campo_arquivo_pdf):
        nome_arquivo = SelectFiles.guardar_nome()
        caminho_arquivo_excel = campo_arquivo_excel.text()
        caminho_arquivo_pdf = campo_arquivo_pdf.text()

        arquivo_excel = openpyxl.load_workbook(caminho_arquivo_excel)
        arquivo_pdf = fitz.open(caminho_arquivo_pdf)

        matriculas_nao_encontradas = list()
        BaseRealcePdf.logger.info('Começando o destaque de PDFs')

        for linha in range(1, arquivo_excel.active.max_row + 1):
            numero_matricula = str(arquivo_excel.active.cell(row=linha, column=2).value)
            nome_funcionario = str(arquivo_excel.active.cell(row=linha, column=3).value).upper()

            encontrou_matricula = False

            if numero_matricula != 'None':
                for pagina in arquivo_pdf:
                    for linha_texto in pagina.get_text().splitlines():
                        if numero_matricula in linha_texto:
                            realce = pagina.search_for(numero_matricula, hit_max=1)
                            if realce:
                                retangulo_realce = fitz.Rect(realce[0][:4])
                                pagina.add_highlight_annot(retangulo_realce)
                                encontrou_matricula = True
                                BaseRealcePdf.logger.info(
                                    f'Destacando a matrícula {numero_matricula} do funcionário {nome_funcionario}')
                                break

            if not encontrou_matricula and nome_funcionario and numero_matricula != "None":
                matriculas_nao_encontradas.append(f"{numero_matricula} - {nome_funcionario}")

        return arquivo_pdf, matriculas_nao_encontradas, nome_arquivo

    @staticmethod
    def salvar_arquivo_pdf(pasta_destino, nome_arquivo, arquivo_pdf):
        numero_arquivo = 1

        while os.path.exists(os.path.join(pasta_destino, nome_arquivo)):
            match = re.search(r'\((\d+)\)', nome_arquivo)

            if match:
                numero_arquivo_existente = int(match.group(1))
                novo_numero_arquivo = numero_arquivo_existente + 1
                nome_arquivo = re.sub(r'\((\d+)\)', f'({novo_numero_arquivo})', nome_arquivo)
            else:
                nome_arquivo = f"{os.path.splitext(nome_arquivo)[0]}({numero_arquivo}).pdf"

            numero_arquivo += 1

        BaseRealcePdf.caminho_arquivo_saida = os.path.join(pasta_destino, nome_arquivo)
        arquivo_pdf.save(BaseRealcePdf.caminho_arquivo_saida)

        return nome_arquivo

    @staticmethod
    def exibir_mensagem_conclusao(pasta_destino, matriculas_nao_encontradas):
        messagebox.showinfo("Concluído", f"O Arquivo final foi salvo em:\n{pasta_destino}")
        BaseRealcePdf.logger.info(f'O Arquivo final foi salvo em: {pasta_destino}')

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
            BaseRealcePdf.logger.info('Algumas matrículas não foram encontradas')


class RealceMatriculas(BaseRealcePdf):

    @staticmethod
    def pdf(pasta_destino, campo_arquivo_excel, campo_arquivo_pdf):
        arquivo_pdf, matriculas_nao_encontradas, nome_arquivo = BaseRealcePdf.destacar_pdf(campo_arquivo_excel,
                                                                                           campo_arquivo_pdf)

        RealceMatriculas.salvar_arquivo_pdf(pasta_destino, nome_arquivo, arquivo_pdf)
        RealceMatriculas.exibir_mensagem_conclusao(pasta_destino, matriculas_nao_encontradas)
