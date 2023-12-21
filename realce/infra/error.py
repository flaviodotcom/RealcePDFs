import os
from tkinter import messagebox


class TratarErro(Exception):
    pass


def existe_erro(caminho_arquivo_excel, caminho_arquivo_pdf):
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
            messagebox.showerror("Erro", mensagem_erro)
            return True
        return False

    except TratarErro as e:
        messagebox.showerror("Erro", f"Ocorreu um erro: {str(e)}")
        return True


def tratar_pasta_destino(pasta_destino):
    if not (pasta_destino and os.path.exists(pasta_destino)):
        messagebox.showerror("Erro", "Selecione uma pasta de destino válida.")
        return True
    return False


def confirmar_diretorio(pasta_destino):
    return messagebox.askokcancel(
        title="Revise as informações",
        message=f"Diretório escolhido:\n{pasta_destino}.\nDeseja Continuar?"
    )
