from realce.app.info import SegundaJanela
from realce.infra.helper import resource_path
from realce.core.vt import SepararPDF
from realce.core.selecionar import SelectFiles
from realce.core.salvar import salvar_para_pasta_padrao, salvar_para_pasta_selecionada_pelo_usuario

from customtkinter import set_appearance_mode, set_default_color_theme, CTk, CTkFrame, CTkLabel, CTkEntry, CTkButton, EW


class HomeWindow:
    root = CTk()

    def __init__(self):
        self.appearance()

    def appearance(self):
        set_appearance_mode("dark")
        set_default_color_theme("dark-blue")

        self.size()
        self.root.title(" Destacar PDFs por matrícula")
        self.root.iconbitmap(resource_path("resources/images/Cookie-Monster.ico"))
        self.root.resizable(False, False)
        self.root.protocol("WM_DELETE_WINDOW", self.fechar)

    def size(self):
        window_height = 220
        window_width = 560

        screen_width = self.root.winfo_screenwidth()
        screen_height = self.root.winfo_screenheight()

        x_cordinate = int((screen_width / 2) - (window_width / 2))
        y_cordinate = int((screen_height / 2) - (window_height / 2))

        self.root.geometry("{}x{}+{}+{}".format(window_width, window_height, x_cordinate, y_cordinate))

    @staticmethod
    def abrir_info():
        info = SegundaJanela()
        info.abrir_janela()

    def fechar(self):
        self.root.destroy()


class HomeWidgets(HomeWindow):
    frame_pdf: CTkFrame
    campo_arquivo_excel: CTkEntry
    campo_arquivo_pdf: CTkEntry

    def __init__(self):
        super().__init__()
        self.build_layout()

    def build_layout(self):
        self.build_campo_excel()
        self.build_campo_pdf()
        self.build_buttons_salvar()

    def build_campo_excel(self):
        frame_excel = CTkFrame(self.root)
        frame_excel.grid(sticky=EW, padx=10, pady=10)

        rotulo_arquivo_excel = CTkLabel(frame_excel, text="  Arquivo Excel:")
        rotulo_arquivo_excel.grid(row=0, column=0, padx=3)

        self.campo_arquivo_excel = CTkEntry(frame_excel, placeholder_text="Selecione o arquivo Excel", width=300)
        self.campo_arquivo_excel.grid(row=0, column=1, padx=5)

        botao_selecionar_arquivo_excel = CTkButton(frame_excel, text="Selecionar",
                                                   command=lambda: SelectFiles.selecionar_arquivo_excel(
                                                       self.campo_arquivo_excel))
        botao_selecionar_arquivo_excel.grid(row=0, column=2)

    def build_campo_pdf(self):
        self.frame_pdf = CTkFrame(self.root)
        self.frame_pdf.grid(sticky=EW, padx=10, pady=10)

        rotulo_arquivo_pdf = CTkLabel(self.frame_pdf, text="  Arquivo PDF:  ")
        rotulo_arquivo_pdf.grid(row=0, column=0, padx=3)

        self.campo_arquivo_pdf = CTkEntry(self.frame_pdf, placeholder_text="Selecione o arquivo PDF", width=300)
        self.campo_arquivo_pdf.grid(row=0, column=1, padx=5)

        botao_selecionar_arquivo_pdf = CTkButton(self.frame_pdf, text="Selecionar",
                                                 command=lambda: SelectFiles.selecionar_arquivo_pdf(
                                                     self.campo_arquivo_pdf))
        botao_selecionar_arquivo_pdf.grid(row=0, column=2)

    def build_buttons_salvar(self):
        frame_salvar_e_info = self.frame_botoes()

        botao_destacar = CTkButton(frame_salvar_e_info, text="Salvar",
                                   command=lambda: salvar_para_pasta_padrao(self.campo_arquivo_excel,
                                                                            self.campo_arquivo_pdf))
        botao_destacar.grid(row=0, column=0, sticky=EW, padx=3)

        botao_destacar_em_outra_pasta = CTkButton(frame_salvar_e_info, text="Salvar Como",
                                                  command=lambda: salvar_para_pasta_selecionada_pelo_usuario(
                                                      self.campo_arquivo_excel, self.campo_arquivo_pdf))
        botao_destacar_em_outra_pasta.grid(row=0, column=1, sticky=EW, padx=3)

        self.button_separar_vts(frame_salvar_e_info)
        self.button_info(frame_salvar_e_info)

    def button_separar_vts(self, frame_salvar_e_info):
        separar_vts = CTkButton(frame_salvar_e_info, text="Separar PDFs por Matrícula",
                                command=lambda: SepararPDF.separar_vt(self.campo_arquivo_excel, self.campo_arquivo_pdf))
        separar_vts.grid(row=1, column=0, sticky=EW, padx=3, pady=10, columnspan=2)
        return separar_vts

    def button_info(self, frame_salvar_e_info):
        botao_expansor_informacoes = CTkButton(frame_salvar_e_info, text="Como funciona?", command=self.abrir_info)
        botao_expansor_informacoes.grid(row=2, column=0, columnspan=2, padx=3, pady=3)

        return botao_expansor_informacoes

    def frame_botoes(self):
        frame_salvar_e_info = CTkFrame(self.root)
        frame_salvar_e_info.grid(row=2, column=0, columnspan=3, sticky=EW, pady=5, padx=10)

        frame_salvar_e_info.columnconfigure(0, weight=1)
        frame_salvar_e_info.columnconfigure(1, weight=1)
        frame_salvar_e_info.configure(fg_color="transparent")

        return frame_salvar_e_info


def carregar_janela(self):
    self.root.mainloop()
