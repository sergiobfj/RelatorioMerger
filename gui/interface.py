# gui/interface.py

import customtkinter as ctk
from tkinter import filedialog
from merger.processor import mesclar_relatorios
from merger.formatter import formatar_excel

def iniciar_gui():
    ctk.set_appearance_mode("System")  # ou "Dark" / "Light"
    ctk.set_default_color_theme("blue")  # padrão do CTk
    
    janela = ctk.CTk()
    janela.title("Mesclagem de Relatórios")
    janela.geometry("500x300")

    path_pos = ""
    path_vendas = ""

    def escolher_pos():
        nonlocal path_pos
        path_pos = filedialog.askopenfilename()

    def escolher_vendas():
        nonlocal path_vendas
        path_vendas = filedialog.askopenfilename()

    def mesclar():
        if not path_pos or not path_vendas:
            print("ERRO: selecione os arquivos!")
            return
        
        output = mesclar_relatorios(path_pos, path_vendas)
        formatar_excel(output)
        print("Relatório Mesclado!")

    # Interface
    label1 = ctk.CTkLabel(janela, text="📂 Arquivo Pós-Venda:")
    label1.pack(pady=10)

    btn1 = ctk.CTkButton(janela, text="Selecionar", command=escolher_pos)
    btn1.pack()

    label2 = ctk.CTkLabel(janela, text="📑 Arquivo VendasI:")
    label2.pack(pady=10)

    btn2 = ctk.CTkButton(janela, text="Selecionar", command=escolher_vendas)
    btn2.pack()

    btn3 = ctk.CTkButton(janela, text="▶ Mesclar", command=mesclar)
    btn3.pack(pady=20)

    janela.mainloop()
