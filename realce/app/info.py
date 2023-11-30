from customtkinter import CTk, CTkLabel, CTkFrame

from realce.infra.helper import resource_path


class SegundaJanela:
    def __init__(self):
        self.sec_window = None
        self.aberta = False

    def abrir_janela(self):
        if not self.aberta:
            self.sec_window = CTk()
            self.sec_window.title("Como Funciona?")
            self.sec_window.resizable(False, False)
            self.sec_window.config(padx=10, pady=10)
            self.sec_window.protocol("WM_DELETE_WINDOW", self.fechar_seg_janela)
            self.sec_window.iconbitmap(resource_path("resources/images/vigarista.ico"))

            painel_informacoes = CTkFrame(self.sec_window)
            painel_informacoes.pack()

            texto_informacoes = (
                "Este programa permite destacar a matrícula dos funcionários em um arquivo PDF usando informações de "
                "uma planilha do Excel."
                "Principalmente usado para realçar os benefícios de Seguro de Vida, Plano Odontológico, "
                "Vale Transporte, Vale Alimentação e Vale Refeição.\n\n"
                "Orientações:\n\n"
                "1. Clique no botão 'Selecionar' para escolher o arquivo Excel e PDF, respectivamente.\n\n"
                "2. Certifique-se de que as matrículas estejam na coluna B da planilha. O programa percorre pela "
                "segunda coluna (coluna B),"
                "garanta que nessa coluna não existam outras informações.\n\n"
                "3. O programa cria um arquivo de texto, que aponta quais foram as matrículas não encontradas junto "
                "com o nome do colaborador."
                "Para que essa funcionalidade ocorra como esperado, mantenha o nome dos colaboradores na terceira "
                "coluna (coluna C) da planilha."
                "O arquivo de texto será salvo no mesmo diretório do PDF editado.\n\n"
                "4. Após selecionar os arquivos, clique no botão 'Salvar' para guardar o PDF editado na Área de "
                "Trabalho (Desktop),"
                "dentro da pasta 'BENEFÍCIOS DESTACADOS'. Alternativamente, clique no botão 'Salvar Como' para "
                "escolher o local de armazenamento"
                "do PDF editado que preferir.\n\n"
                "Separar PDFS:\n\n"
                "1. Clique no botão 'Selecionar' para escolher o arquivo Excel e PDF, respectivamente.\n\n"
                "2. Após selecionar os arquivos, clique no botão 'Separar PDFs por Matricula' para escolher a pasta "
                "de destino do PDF realçado,"
                "logo após o programa irá criar uma pasta chamada 'Vt Separados', onde irá conter cada pagina do PDF "
                "realçado que encontrou a"
                "matrícula, e no fim irá juntar todos os PDFs que estão na pasta em um arquivo chamado 'Vt Final'."
            )

            tamanho_da_fonte = 14
            fonte_personalizada = ("Helvetica", tamanho_da_fonte)

            rotulo_informacoes = CTkLabel(
                painel_informacoes,
                text=texto_informacoes,
                wraplength=505,
                justify="left",
                font=fonte_personalizada,
            )
            rotulo_informacoes.pack(pady=10, padx=10)
            rotulo_informacoes.configure(text_color="white")

            self.aberta = True
            self.sec_window.mainloop()

    def fechar_seg_janela(self):
        if self.aberta:
            self.sec_window.destroy()
            self.aberta = False


if __name__ == '__main__':
    run = SegundaJanela()
    run.abrir_janela()
