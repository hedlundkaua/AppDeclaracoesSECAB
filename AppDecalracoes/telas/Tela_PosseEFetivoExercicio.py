import os
import locale
import tkinter as tk
import formatarTexto as fd
from tkinter import messagebox
from datetime import datetime
from docxtpl import DocxTemplate
from tkinter import ttk

def abrir_tela():

    locale.setlocale(locale.LC_TIME, "pt_BR.UTF-8")
    # pega o texto sem barras, até 8 caracteres

    def gerar_documento():
        try:
            nome = entry_nome.get()
            id = entry_id.get()
            cpf = entry_cpf.get()
            rg = entry_rg.get()
            cargoAtual = entry_cargo.get()
            dataDaPublicacao = entry_publicacao.get()
            pagina = entry_pagina.get()
            dataInicioExe = entry_exercicio.get()
            genero = genero_var.get()
            servidor_assinador = combo_servidor.get()
            
            #caminho do modelo
            modelo_path = r"W:\DRH\SECAB\Kaua Teste\modelos\Declaracao_PosseEEfetivoExercicio.docx"

            #caminho de saida
            saida_path = r"W:\DRH\SECAB\Kaua Teste\gerados"

            if not os.path.exists(modelo_path):
                messagebox.showerror("Error", f"Caminho não encontrado \n{modelo_path}")
                return
            
            doc = DocxTemplate(modelo_path)


            data_arquivo = datetime.now().strftime("%d de %B de %Y")
            
            dados_assinador = servidores.get(servidor_assinador, {})

            contexto = {
                "nome": nome,
                "cpf": cpf,
                "rg": rg,
                "id": id,
                "cargoAtual": cargoAtual,
                "dataDaPublicacao": dataDaPublicacao,
                "dataInicioExercicio": dataInicioExe,
                "pagina": pagina,
                "data": data_arquivo,
                "genero": genero
            }

            data_hoje = datetime.now().strftime("%d-%M-%Y")
            nome_arquivo = f"{nome.replace(' ','_')}_declaracao_PosseEEfetivoExercicio {data_hoje}.docx"
            caminho_saida = os.path.join(saida_path, nome_arquivo)


            contexto.update(dados_assinador)
            doc.render(contexto) # executa o contexto substituindo os campos do word para os que o usuario informa no app
            doc.save(caminho_saida) #executa e salva o arquivo com as modificações no caminho de saida escolhido

            messagebox.showinfo("Sucesso", f"Documento gerado com sucesso:\n{caminho_saida}")
            janela.destroy()

            if not nome or not id or not cpf or not cargoAtual or not pagina or not rg or not dataDaPublicacao or not dataInicioExe:
                messagebox.showwarning("Campos obrigatórios, por favor preencha todos os campos!")
                return

        except Exception as e:
            messagebox.showerror("Erro", f"Ocorreu um erro:\n{e}")

    janela = tk.Tk()
    janela.title("Declaração de Posse e Efetivo Exercicio")

    tk.Label(janela, text="Nome: ").grid(row=0, column=0, sticky="e")
    entry_nome = tk.Entry(janela, width=40)
    entry_nome.grid(row=0, column=1)

    tk.Label(janela, text="ID Funcional: ").grid(row=1, column=0, sticky="e")
    entry_id = tk.Entry(janela, width=40)
    entry_id.grid(row=1, column=1)

    tk.Label(janela, text="RG: ").grid(row=2, column=0, sticky="e")
    entry_rg = tk.Entry(janela, width=40)
    entry_rg.grid(row=2, column=1)

    tk.Label(janela, text="CPF: ").grid(row=3, column=0, sticky="e")
    entry_cpf = tk.Entry(janela, width=40)
    entry_cpf.grid(row=3, column=1)
    entry_cpf.bind("<KeyRelease>", lambda event: fd.formatar_cpf(event, entry_cpf))

    tk.Label(janela, text="Cargo Atual:").grid(row=4, column=0, sticky="e")
    entry_cargo = tk.Entry(janela, width=40)
    entry_cargo.grid(row=4, column=1)

    # Data da publicação
    tk.Label(janela, text="Data da publicação:").grid(row=5, column=0, sticky="e")
    entry_publicacao = tk.Entry(janela, width=40) #onde o usuario vai digitar
    entry_publicacao.grid(row=5, column=1) 
    # o bind("<KeyRelease>",) significa que toda vez que o usuario soltar uma tecla dentro desse campo vai chamar a função formatar_data
    entry_publicacao.bind("<KeyRelease>", lambda event: fd.formatar_data(event, entry_publicacao))
    #passamos o proprio evento do teclado(event) e a referentcia do campo para a função.
    #event = objeto que representa o evento do teclado que acabou de acontecer.
    #assim, a função pega o texto digitado até o momento e aplica a formatação automaticamente.

    tk.Label(janela, text="Pagina:").grid(row=6, column=0, sticky="e")
    entry_pagina = tk.Entry(janela, width=40)
    entry_pagina.grid(row=6, column=1)

    tk.Label(janela, text="Inicio do Exercicio:").grid(row=7, column=0, sticky="e")
    entry_exercicio = tk.Entry(janela, width=40)
    entry_exercicio.grid(row=7, column=1)
    entry_exercicio.bind("<KeyRelease>", lambda event: fd.formatar_data(event, entry_exercicio))

    
    genero_var = tk.StringVar(value="Masculino")
    tk.Label(janela, text="Genero:").grid(row=8, column=0, sticky="e")
    tk.Radiobutton(janela, text="Masculino", variable=genero_var, value="Masculino").grid(row=8, column=1, sticky="w")
    tk.Radiobutton(janela, text="Feminino", variable=genero_var, value="Feminino").grid(row=9, column=1, sticky="w")

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

    tk.Label(janela, text="Quem vai assinar:").grid(row=9, column=0, sticky="e")
    combo_servidor = ttk.Combobox(janela, values=list(servidores.keys()), width= 37)
    combo_servidor.grid(row=9, column=1)


    bnt_gerar = tk.Button(janela, text="Gerar Documento", command=gerar_documento)
    bnt_gerar.grid(row=12, column=0, columnspan=2, pady=10)

    janela.mainloop()