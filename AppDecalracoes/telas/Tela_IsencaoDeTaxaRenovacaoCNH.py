import os
import locale
import tkinter as tk
import formatarTexto as fd
from tkinter import messagebox
from datetime import datetime
from docxtpl import DocxTemplate


def abrir_tela():

    locale.setlocale(locale.LC_TIME, "pt_BR.UTF-8")
    # pega o texto sem barras, até 8 caracteres

    def gerar_documento():
        try:
            nome = entry_nome.get()
            id = entry_id.get()
            cpf = entry_cpf.get()
            classe = entry_classe.get()
            rg = entry_rg.get()
            pai = entry_pai.get()
            mae = entry_mae.get()
            cidade = entry_cidade.get()
            dataNascimento = entry_nascimento.get()
            genero = genero_var.get()

            #caminho do modelo
            modelo_path = r"W:\DRH\SECAB\Kaua Teste\modelos\declaracao_IsencaoDeTaxaRenovacaoDeCNH.docx"

            #caminho de saida
            saida_path = r"W:\DRH\SECAB\Kaua Teste\gerados"

            if not os.path.exists(modelo_path):
                messagebox.showerror("Error", f"Caminho não encontrado \n{modelo_path}")
                return
            
            doc = DocxTemplate(modelo_path)


            data_arquivo = datetime.now().strftime("%d de %B de %Y")

            contexto = {
                "nome": nome,
                "cpf": cpf,
                "id": id,
                "classe": classe,
                "pai": pai,
                "mae": mae,
                "rg": rg,
                "cidade": cidade,
                "dataNascimento": dataNascimento,
                "classe": classe,
                "data": data_arquivo,
                "genero": genero
            }

            data_hoje = datetime.now().strftime("%d-%M-%Y")
            nome_arquivo = f"{nome.replace(' ','_')}_declaracao_IsencaoCNH {data_hoje}.docx"
            caminho_saida = os.path.join(saida_path, nome_arquivo)

            doc.render(contexto) # executa o contexto substituindo os campos do word para os que o usuario informa no app
            doc.save(caminho_saida) #executa e salva o arquivo com as modificações no caminho de saida escolhido

            messagebox.showinfo("Sucesso", f"Documento gerado com sucesso:\n{caminho_saida}")
            janela.destroy()
 
            if (not nome or not id or not cpf or not classe or not pai or not mae or not rg or not dataNascimento):
              messagebox.showwarning("Campos obrigatórios, por favor preencha todos os campos!")
              return

        except Exception as e:
            messagebox.showerror("Erro", f"Ocorreu um erro:\n{e}")

    janela = tk.Tk()
    janela.title("Delcaração d Vínculo")

    tk.Label(janela, text="Nome: ").grid(row=0, column=0, sticky="e")
    entry_nome = tk.Entry(janela, width=40)
    entry_nome.grid(row=0, column=1)

    tk.Label(janela, text="ID Funcional: ").grid(row=1, column=0, sticky="e")
    entry_id = tk.Entry(janela, width=40)
    entry_id.grid(row=1, column=1)

    tk.Label(janela, text="CPF: ").grid(row=2, column=0, sticky="e")
    entry_cpf = tk.Entry(janela, width=40)
    entry_cpf.grid(row=2, column=1)
    entry_cpf.bind("<KeyRelease>", lambda event: fd.formatar_cpf(event, entry_cpf))

    tk.Label(janela, text="RG: ").grid(row=3, column=0, sticky="e")
    entry_rg = tk.Entry(janela, width=40)
    entry_rg.grid(row=3, column=1)
    entry_rg.bind("<KeyRelease>", lambda event: fd.formatar_cpf(event, entry_cpf))

    tk.Label(janela, text="Data de Nascimneto:").grid(row=4, column=0, sticky="e")
    entry_nascimento = tk.Entry(janela, width=40)
    entry_nascimento.grid(row=4, column=1)
    entry_nascimento.bind("<KeyRelease>", lambda event: fd.formatar_data(event, entry_nascimento))

    tk.Label(janela, text="Classe:").grid(row=5, column=0, sticky="e")
    entry_classe = tk.Entry(janela, width=40)
    entry_classe.grid(row=5, column=1)

    tk.Label(janela, text="Cidade: ").grid(row=6, column=0, sticky="e")
    entry_cidade = tk.Entry(janela, width=40)
    entry_cidade.grid(row=6, column=1)
    
    tk.Label(janela, text="Nome da Mãe:").grid(row=7, column=0, sticky="e")
    entry_mae = tk.Entry(janela, width=40)
    entry_mae.grid(row=7, column=1)

    tk.Label(janela, text="Nome de Pai:").grid(row=8, column=0, sticky="e")
    entry_pai = tk.Entry(janela, width=40)
    entry_pai.grid(row=8, column=1)


    # genero com radiusButtons
    genero_var = tk.StringVar(value="Masculino") # valor padrão
    tk.Label(janela, text="Genero:").grid(row=11, column=0, sticky="e")
    tk.Radiobutton(janela, text="Masculino", variable=genero_var, value="Masculino").grid(row=11, column=1, sticky="w")
    tk.Radiobutton(janela, text="Feminino", variable=genero_var, value="Feminino").grid(row=12, column=1, sticky="w")


    bnt_gerar = tk.Button(janela, text="Gerar Documento", command=gerar_documento)
    bnt_gerar.grid(row=14, column=0, columnspan=2, pady=10)

    janela.mainloop()