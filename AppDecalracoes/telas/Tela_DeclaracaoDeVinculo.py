import locale
import tkinter as tk
from tkinter import messagebox
from datetime import datetime
from docxtpl import DocxTemplate
import os


def abrir_tela():

    locale.setlocale(locale.LC_TIME, "pt_BR.UTF-8")

    def gerar_documento():
        try:
            nome = entry_nome.get()
            id = entry_id.get()
            cargoAtual = entry_cargo.get()
            classe = entry_classe.get()
            dataDaPublicacao = entry_publicacao.get()
            pagina = entry_pagina.get()
            dataInicioExe = entry_exercicio.get()
            setorAtual = entry_setor.get()

            #caminho do modelo
            modelo_path = r"W:\DRH\SECAB\Kaua Teste\modelos\declaracao_vinculo.docx"

            #caminho de saida
            saida_path = r"W:\DRH\SECAB\Kaua Teste\gerados"

            if not os.path.exists(modelo_path):
                messagebox.showerror("Error", f"Caminho não encontrado \n{modelo_path}")
                return
            
            doc = DocxTemplate(modelo_path)


            data_arquivo = datetime.now().strftime("%d de %B de %Y")

            contexto = {
                "nome": nome,
                "id": id,
                "cargoAtual": cargoAtual,
                "classe": classe,
                "dataDaPublicacao": dataDaPublicacao,
                "pagina": pagina,
                "dataInicioExercicio": dataInicioExe,
                "setorAtual": setorAtual,
                "classe": classe,
                "data": data_arquivo 
            }

            data_hoje = datetime.now().strftime("%d-%M-%Y")
            nome_arquivo = f"{nome.replace(' ','_')}_declaracao_vinculo {data_hoje}.docx"
            caminho_saida = os.path.join(saida_path, nome_arquivo)

            doc.render(contexto) # executa o contexto substituindo os campos do word para os que o usuario informa no app
            doc.save(caminho_saida) #executa e salva o arquivo com as modificações no caminho de saida escolhido

            messagebox.showinfo("Sucesso", f"Documento gerado com sucesso:\n{caminho_saida}")
            janela.destroy()

            if not nome or not id or not cargoAtual or not classe or not dataDaPublicacao or not pagina or not dataInicioExe or not setorAtual:
                messagebox.showwarning("Campos obrigatórios, por favor preencha todos os campos!")
                return

        except Exception as e:
            messagebox.showerror("Erro", f"Ocorreu um erro:\n{e}")

    janela = tk.Tk()
    janela.title("Delcaração d Vínculo")

    tk.Label(janela, text="Nome:").grid(row=0, column=0, sticky="e")
    entry_nome = tk.Entry(janela, width=40)
    entry_nome.grid(row=0, column=1)

    tk.Label(janela, text="ID Funcional:").grid(row=1, column=0, sticky="e")
    entry_id = tk.Entry(janela, width=40)
    entry_id.grid(row=1, column=1)

    tk.Label(janela, text="Cargo Atual:").grid(row=2, column=0, sticky="e")
    entry_cargo = tk.Entry(janela, width=40)
    entry_cargo.grid(row=2, column=1)

    tk.Label(janela, text="Classe:").grid(row=3, column=0, sticky="e")
    entry_classe = tk.Entry(janela, width=40)
    entry_classe.grid(row=3, column=1)

    tk.Label(janela, text="Data da publicação:").grid(row=4, column=0, sticky="e")
    entry_publicacao = tk.Entry(janela, width=40)
    entry_publicacao.grid(row=4, column=1)

    tk.Label(janela, text="Página:").grid(row=5, column=0, sticky="e")
    entry_pagina = tk.Entry(janela, width=40)
    entry_pagina.grid(row=5, column=1)

    tk.Label(janela, text="Inicio do Exercicio:").grid(row=6, column=0, sticky="e")
    entry_exercicio = tk.Entry(janela, width=40)
    entry_exercicio.grid(row=6, column=1)

    tk.Label(janela, text="Setor Atual:").grid(row=7, column=0, sticky="e")
    entry_setor = tk.Entry(janela, width=40)
    entry_setor.grid(row=7, column=1)

    bnt_gerar = tk.Button(janela, text="Gerar Documento", command=gerar_documento)
    bnt_gerar.grid(row=8, column=0, columnspan=2, pady=10)

    janela.mainloop()