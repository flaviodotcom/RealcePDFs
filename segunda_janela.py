import customtkinter as ctk


class SegundaJanela():
    def __init__(self):
        self.sec_window = None
        
    def abrir_janela(self):
        if self.sec_window is None:
            self.sec_window = ctk.CTk()
            self.sec_window.title("Como Funciona?")
            self.sec_window.resizable(False, False)
            self.sec_window.config(padx=10, pady=10)
            
            painel_informacoes = ctk.CTkFrame(self.sec_window)
            painel_informacoes.pack()
            
            texto_informacoes = """
            Este programa permite destacar a matrícula dos funcionários em um arquivo PDF usando informações de uma planilha do Excel. Principalmente usado para realçar os benefícios de Seguro de Vida, Plano Odontológico, Vale Transporte, Vale Alimentação e Vale Refeição.

                                                            Orientações:

            1. Clique no botão 'Selecionar' para escolher o arquivo Excel e PDF, respectivamente.

            2. Certifique-se de que as matrículas estejam na coluna B da planilha. O programa percorre pela segunda coluna (coluna B), garanta que nesse coluna não existam outras informações.

            3. O programa cria um arquivo de texto, que aponta quais foram as matrículas não encontradas junto com o nome do colaborador. Para que essa funcionalidade ocorra como esperado, mantenha o nome dos colaboradores na terceira coluna (coluna C) da planilha. O arquivo de texto será salvo no mesmo diretório do PDF editado.

            4. Após selecionar os arquivos, clique no botão 'Salvar' para guardar o PDF editado na Área de Trabalho (Desktop), dentro da pasta 'BENEFÍCIOS DESTACADOS'. Alternativamente, clique no botão 'Salvar Como' para escolher o local de armazenamento do PDF editado que preferir.
            """

            tamanho_da_fonte = 14
            fonte_personalizada = ("Helvetica", tamanho_da_fonte)
            
            rotulo_informacoes = ctk.CTkLabel(
                painel_informacoes,
                text=texto_informacoes,
                wraplength=505,
                justify="left",
                font=fonte_personalizada,
            )
            rotulo_informacoes.pack(pady=10, padx=10)
            rotulo_informacoes.configure(text_color="white")

            self.sec_window.mainloop()
            