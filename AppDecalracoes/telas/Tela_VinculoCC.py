import os
import tkinter as tk
from tkinter import messagebox
from docxtpl import DocxTemplate
from datetime import datetime
import formatarData as fd
import locale

locale.setlocale(locale.LC_TIME, "pt_BR.UTF-8")

def abrir_tela():
    def gerar_documento():
        try:
            nome = entry_nome.get()
            id = entry_id.get()
            cargoAtual = entry_cargoAtual.get()
            dataPublicacao = entry_dataPublicacao.get()
            pagina = entry_pagina.get()
            dataExe = entry_dataExe.get()
            setorAtual = entry_setorAtual.get()
            genero = genero_var.get()


            caminho_modelo = r"W:\DRH\SECAB\Kaua Teste\modelos\declaracao_vinculo_cc.docx"
            saida_arquivo = r"W:\DRH\SECAB\Kaua Teste\gerados"

            if not os.path.exists(caminho_modelo):
                messagebox.showerror("Erro", f"Caminho não encontrado zn{caminho_modelo}")
                return
            
            doc = DocxTemplate(caminho_modelo)

            data_arquivo = datetime.now().strftime("%d de %B de %Y")

            informacoes = {
                "nome": nome,
                "id": id,
                "cargoAtual": cargoAtual,
                "dataDaPublicacao": dataPublicacao,
                "pagina": pagina,
                "dataInicioDeExercicio": dataExe,
                "setorAtual": setorAtual,
                "data": data_arquivo,
                "genero": genero
            }

            data_hoje = datetime.now().strftime("%d-%M-%Y")
            nome_arquivo = f"{nome.replace(' ', '_')}_declaracao_vinculo {data_hoje}.docx"
            caminho_saida = os.path.join(saida_arquivo, nome_arquivo)

            doc.render(informacoes)
            doc.save(caminho_saida)

            messagebox.showinfo("Sucesso", f"Documento gerado com sucesso:\n{caminho_saida}")
            janela.destroy()

            if (not nome or not id or not cargoAtual or not dataPublicacao 
                or not pagina or not dataExe or not setorAtual):
                messagebox.showwarning("Campos obrigatorios , por favor preencha todos os campos!") 
                return
        except Exception as e:
            messagebox.showerror("Erro", f"Ocorreu um erro:\n{e}")
            return 

    janela = tk.Tk()
    janela.title("Declaração de Vinculo CC")

    tk.Label(janela, text="Nome: ").grid(row=0, column=0, sticky="e")
    entry_nome = tk.Entry(janela, width=40)
    entry_nome.grid(row=0, column=1)

    tk.Label(janela, text="id: ").grid(row=1, column=0, sticky="e")
    entry_id = tk.Entry(janela, width=40)
    entry_id.grid(row=1, column=1)

    tk.Label(janela, text="Cargo Atual: ").grid(row=2, column=0, sticky="e")
    entry_cargoAtual = tk.Entry(janela, width=40)
    entry_cargoAtual.grid(row=2, column=1)

    tk.Label(janela, text="Data da Publicação: ").grid(row=3, column=0, sticky="e")
    entry_dataPublicacao = tk.Entry(janela, width=40)
    entry_dataPublicacao.grid(row=3, column=1)
    entry_dataPublicacao.bind("<KeyRelease>", lambda event: fd.formatar_data(event, entry_dataPublicacao))

    tk.Label(janela, text="Página: ").grid(row=4, column=0, sticky="e")
    entry_pagina = tk.Entry(janela, width=40)
    entry_pagina.grid(row=4, column=1)

    tk.Label(janela, text="Data de Exercício: ").grid(row=5, column=0, sticky="e")
    entry_dataExe = tk.Entry(janela, width=40)
    entry_dataExe.grid(row=5, column=1)
    entry_dataExe.bind("<KeyRelease>", lambda event: fd.formatar_data(event, entry_dataExe))

    tk.Label(janela, text="Setor Atual: ").grid(row=6, column=0, sticky="e")
    entry_setorAtual = tk.Entry(janela, width=40)
    entry_setorAtual.grid(row=6, column=1)

    genero_var = tk.StringVar(value="Masculino")
    tk.Label(janela, text="Genero:").grid(row=7, column=0, sticky="e")
    tk.Radiobutton(janela, text="Masculino", variable=genero_var, value="Masculino").grid(row=7, column=1, sticky="w")
    tk.Radiobutton(janela, text="Feminino", variable=genero_var, value="Feminino").grid(row=8, column=1, sticky="w")

    bnt_gerar = tk.Button(janela, text="Gerar documento", command=gerar_documento)
    bnt_gerar.grid(row=10, column=0, columnspan=2, pady=10)

    janela.mainloop()