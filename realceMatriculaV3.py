import os
import fitz
import openpyxl
import customtkinter as ctk
from tkinter import filedialog, messagebox

file_name = ""

def select_excel_file():
    file_path = filedialog.askopenfilename(filetypes=[("Excel Files", "*.xlsx"), ("CSV Files", "*.csv"), ("Excel 97-2003 Files", "*.xls")])
    excel_file_entry.delete(0, ctk.END)
    excel_file_entry.insert(0, file_path)

def select_pdf_file():
    file_path = filedialog.askopenfilename(filetypes=[("PDF Files", "*.pdf")])
    global file_name
    pdf_file_entry.delete(0, ctk.END)
    pdf_file_entry.insert(0, file_path)
    file_name = os.path.basename(file_path)

def toggle_info_panel():
    if info_frame.grid_info():
        info_frame.grid_remove()
    else:
        info_frame.grid(row=3, column=0, columnspan=3, padx=10, pady=10, sticky=ctk.EW)

def highlight_registration_numbers(destination_folder):
    global file_name
    pdf_file_path = pdf_file_entry.get()
    if not pdf_file_path:
        ctk.messagebox.showerror("Erro", "Por favor, selecione o arquivo PDF.")
        return

    pdf_file = fitz.open(pdf_file_path)

    excel_file_path = excel_file_entry.get()
    if not excel_file_path:
        ctk.messagebox.showerror("Erro", "Por favor, selecione o arquivo Excel.")
        return

    excel_file = openpyxl.load_workbook(excel_file_path)
    worksheet = excel_file.active

    num_rows = worksheet.max_row

    for row in range(1, num_rows + 1):
        registration_number = worksheet.cell(row=row, column=2).value
        registration_number = str(registration_number)

        for page in pdf_file:
            for line in page.get_text().splitlines():
                if registration_number in line:
                    highlight = page.search_for(registration_number, hit_max=1)
                    if highlight:
                        highlight_rect = fitz.Rect(highlight[0][:4])
                        highlight_annot = page.add_highlight_annot(highlight_rect)

    output_filename = file_name
    file_number = 1
    while os.path.exists(os.path.join(destination_folder, output_filename)):
        output_filename = f"{file_name.strip('.pdf')}({file_number}).pdf"
        file_number += 1

    output_file_path = os.path.join(destination_folder, output_filename)
    pdf_file.save(output_file_path)

    messagebox.showinfo("Concluído", f"O PDF editado foi salvo em:\n{output_file_path}")


def save_to_default_folder():
    destination_folder = os.path.join(os.path.dirname(__file__), "BENEFICIOS DESTACADOS")
    os.makedirs(destination_folder, exist_ok=True)
    highlight_registration_numbers(destination_folder)

def save_to_user_selected_folder():
    destination_folder = filedialog.askdirectory()
    if not destination_folder:
        return
    highlight_registration_numbers(destination_folder)

ctk.set_appearance_mode("dark")
ctk.set_default_color_theme("dark-blue")

root = ctk.CTk()
root.title("Destacar PDFs por matrícula")
root.resizable(False, False)

root.columnconfigure(0, weight=1)
root.rowconfigure(1, weight=1)

excel_frame = ctk.CTkFrame(root)
excel_frame.grid(sticky=ctk.EW, padx=10, pady=10)

excel_file_label = ctk.CTkLabel(excel_frame, text="  Arquivo Excel:")
excel_file_label.grid(row=0, column=0, padx=5)

excel_file_entry = ctk.CTkEntry(excel_frame, placeholder_text="Selecione o arquivo Excel", width=300)
excel_file_entry.grid(row=0, column=1, padx=10)

excel_file_button = ctk.CTkButton(excel_frame, text="Selecionar", command=select_excel_file)
excel_file_button.grid(row=0, column=2, padx=5)

pdf_frame = ctk.CTkFrame(root)
pdf_frame.grid(sticky=ctk.EW, padx=10, pady=10)

pdf_file_label = ctk.CTkLabel(pdf_frame, text="  Arquivo PDF:  ")
pdf_file_label.grid(row=0, column=0, padx=5)

pdf_file_entry = ctk.CTkEntry(pdf_frame, placeholder_text="Selecione o arquivo PDF", width=300)
pdf_file_entry.grid(row=0, column=1, padx=10)

pdf_file_button = ctk.CTkButton(pdf_frame, text="Selecionar", command=select_pdf_file)
pdf_file_button.grid(row=0, column=2, padx=5)

save_and_info_frame = ctk.CTkFrame(root)
save_and_info_frame.grid(row=2, column=0, columnspan=3, sticky=ctk.EW, pady=5, padx=10)

highlight_button = ctk.CTkButton(save_and_info_frame, text="Salvar", command=save_to_default_folder)
highlight_button.grid(row=0, column=0, padx=5, sticky=ctk.EW)

highlight_to_other_directory = ctk.CTkButton(save_and_info_frame, text="Salvar Como", command=save_to_user_selected_folder)
highlight_to_other_directory.grid(row=0, column=1, padx=5, sticky=ctk.EW)

save_and_info_frame.columnconfigure(0, weight=1)
save_and_info_frame.columnconfigure(1, weight=1)

info_frame = ctk.CTkFrame(root)
info_frame.columnconfigure(0, weight=1)

info_text = """
Este programa permite destacar a matrícula dos funcionários em um arquivo PDF usando informações de uma planilha do Excel. Principalmente usado para realçar os benefícios de Seguro de Vida, Plano Odontológico, Vale Transporte, Vale Alimentação e Vale Refeição.

Orientações:

1. Clique no botão 'Selecionar' para escolher o arquivo Excel e PDF, respectivamente.

2. Certifique-se de que as matrículas estejam na coluna B da planilha. O programa percorre pela segunda coluna (coluna B), garanta que nesse coluna não existam outras informações.

3. Depois de selecionar os arquivos, clique no botão 'Salvar' para salvar o PDF editado no mesmo diretório em que o programa está localizado, dentro de uma pasta que será criada, chamada 'BENEFICIOS DESTACADOS'. Ou clique no botão 'Salvar Como', para salvar o PDF editado no caminho que preferir.
"""

tamanho_da_fonte = 14
fonte_personalizada = ("Arial", tamanho_da_fonte)

info_label = ctk.CTkLabel(info_frame, text=info_text, wraplength=520, justify="left", font=fonte_personalizada)
info_label.grid(row=0, column=0, padx=10, pady=5)
info_label.configure(text_color="white")

center_frame = ctk.CTkFrame(root)
center_frame.grid(row=4, column=0, columnspan=3, sticky=ctk.EW)
center_frame.configure(fg_color="transparent")

info_expander = ctk.CTkButton(center_frame, text="Como funciona?", command=toggle_info_panel)
info_expander.pack(pady=5)

root.mainloop()
