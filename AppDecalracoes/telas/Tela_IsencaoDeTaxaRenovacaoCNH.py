import os
import locale
import tkinter as tk
import formatarTexto as fd
from tkinter import messagebox
from datetime import datetime
from docxtpl import DocxTemplate
from tkinter import ttk
import sys

locale.setlocale(locale.LC_TIME, "pt_BR.UTF-8")

def caminho_arquivo(nome):
    """Retorna caminho correto do arquivo, seja no .py ou no .exe"""
    if hasattr(sys, '_MEIPASS'):
        return os.path.join(sys._MEIPASS, nome)
    return os.path.join(os.path.abspath("."), nome)


def abrir_tela():
    def gerar_documento():
        try:
            nome = entry_nome.get()
            id = entry_id.get()
            cpf = entry_cpf.get()
            nOficio = entry_nOficio.get()
            classe = entry_classe.get()
            rg = entry_rg.get()
            pai = entry_pai.get()
            mae = entry_mae.get()
            cidade = entry_cidade.get()
            dataNascimento = entry_nascimento.get()
            genero = genero_var.get()
            servidor_assinador = combo_servidor.get()

            #caminho do modelo
            modelo_path = caminho_arquivo("W:\DRH\SECAB\Kaua Teste\modelos\declaracao_IsencaoDeTaxaRenovacaoDeCNH.docx")

            #caminho de saida
            saida_path = r"W:\DRH\SECAB\Kaua Teste\gerados"

            if not os.path.exists(modelo_path):
                messagebox.showerror("Error", f"Caminho não encontrado \n{modelo_path}")
                return
            
            doc = DocxTemplate(modelo_path)


            data_arquivo = datetime.now().strftime("%d de %B de %Y")

            dados_assinador = servidores.get(servidor_assinador, {})

            contexto = {
                "nOficio": nOficio,
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

            contexto.update(dados_assinador)
            doc.render(contexto) # executa o contexto substituindo os campos do word para os que o usuario informa no app
            doc.save(caminho_saida) #executa e salva o arquivo com as modificações no caminho de saida escolhido

            messagebox.showinfo("Sucesso", f"Documento gerado com sucesso:\n{caminho_saida}")
            janela.destroy()
 
            if (not nome or not id or not cpf or not classe or not pai or not mae or not rg or not dataNascimento or not nOficio):
              messagebox.showwarning("Campos obrigatórios, por favor preencha todos os campos!")
              return

        except Exception as e:
            messagebox.showerror("Erro", f"Ocorreu um erro:\n{e}")

    janela = tk.Tk()
    janela.title("Declaração de isenção de taxa renovação CNH")

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

    tk.Label(janela, text="N° Oficio:").grid(row=9, column=0, sticky="e")
    entry_nOficio = tk.Entry(janela, width=40)
    entry_nOficio.grid(row=9, column=1)

    # genero com radiusButtons
    genero_var = tk.StringVar(value="Masculino") # valor padrão
    tk.Label(janela, text="Genero:").grid(row=11, column=0, sticky="e")
    tk.Radiobutton(janela, text="Masculino", variable=genero_var, value="Masculino").grid(row=11, column=1, sticky="w")
    tk.Radiobutton(janela, text="Feminino", variable=genero_var, value="Feminino").grid(row=12, column=1, sticky="w")

    servidores = {
    "Barbara": {
        "nomeAssinador": "Barbara Lopes de Almeida",
        "cargoAssinador": "Analista Tributario da Receita Estadual",
        "classeAssinador": "A",
        "idAssinador": "5047200/01",
        "dataPorExtenso": datetime.now().strftime("%d de %B de %Y")
    },
    "Juiane": {
        "nomeAssinador": "Juiane Da Silva Machado",
        "cargoAssinador": "Analista Tributario da Receita Estadual",
        "classeAssinador": "D",
        "idAssinador": "4349660/01",
        "dataPorExtenso": datetime.now().strftime("%d de %B de %Y")
    }
    }

    tk.Label(janela, text="Quem vai assinar:").grid(row=13, column=0, sticky="e")
    combo_servidor = ttk.Combobox(janela, values=list(servidores.keys()), width= 37)
    combo_servidor.grid(row=13, column=1)

    bnt_gerar = tk.Button(janela, text="Gerar Documento", command=gerar_documento)
    bnt_gerar.grid(row=14, column=0, columnspan=2, pady=10)

    janela.mainloop()