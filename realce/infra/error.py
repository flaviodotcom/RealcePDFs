import os
from tkinter import messagebox

from realce import get_logger


class TratarErro(Exception):
    def __init__(self, message):
        self.logger = get_logger(self.__class__.__name__)
        self.logger.error(message)


def existe_erro(caminho_arquivo_excel, caminho_arquivo_pdf) -> bool:
    try:
        mensagem_erro = (
            "Por favor, selecione o arquivo Excel e o arquivo PDF."
            if not (caminho_arquivo_excel and caminho_arquivo_pdf)
            else "Por favor, verifique o caminho do arquivo PDF."
            if not (caminho_arquivo_pdf and os.path.isfile(caminho_arquivo_pdf))
            else "Por favor, verifique o caminho do arquivo Excel."
            if not (caminho_arquivo_excel and os.path.isfile(caminho_arquivo_excel))
            else None
        )

        if mensagem_erro:
            raise TratarErro(mensagem_erro)
        return False

    except TratarErro as e:
        messagebox.showerror('Ocorreu um problema', str(e))
        return True


def tratar_pasta_destino(pasta_destino) -> bool:
    try:
        if not (pasta_destino and os.path.exists(pasta_destino)):
            raise TratarErro('Selecione uma pasta de destino válida.')
    except TratarErro as e:
        messagebox.showerror('Ocorreu um problema', str(e))
        return True

    return False


def confirmar_diretorio(pasta_destino):
    return messagebox.askokcancel(
        title="Revise as informações",
        message=f"Diretório escolhido:\n{pasta_destino}.\nDeseja Continuar?"
    )
