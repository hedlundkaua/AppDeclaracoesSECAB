import tkinter as tk
from tkinter import messagebox
from datetime import datetime
from docxtpl import DocxTemplate
import os

def abrir_tela():

    def formatar_data(event, entry):
        texto = entry.get().replace("/", "")[:8]  # remove barras e limita 8 dígitos
        formatado = ""

        for i, c in enumerate(texto):
            if i == 2 or i == 4:
                formatado += "/"
            formatado += c

        entry.delete(0, tk.END)
        entry.insert(0, formatado)

    def gerar_documento():
        try:
            nome = entry_nome.get() # entry é o objeto que representa a caixa que o usuario vai digitar e guardar na variavel 'nome' que iniciamos 
            id_funcional = entry_id.get() # get() representa o texto digitado na caixa 
            data_doe = entry_doe.get() 
            data_exe = entry_exe.get()
            cargo = entry_cargo.get()
            genero = genero_var.get()
            
            if not nome or not id_funcional or not data_doe or not data_exe or not cargo:
                messagebox.showwarning("Campos obrigatórios", "Por favor, preencha todos os campos.")
                return

            # Caminho do modelo
            modelo_path = r"W:\DRH\SECAB\Kaua Teste\modelos\declaracao_inss.docx"

            #caminho da saida
            saida_path = r"W:\DRH\SECAB\Kaua Teste\gerados"


            if not os.path.exists(modelo_path): #o modulo 'os' verifica se no caminho especificado em modelo_path o arquivo ou pasta existe 
                messagebox.showerror("Erro", f"Modelo não encontrado:\n{modelo_path}") #se der verdadeiro executa a mensagem de erro 
                return

            doc = DocxTemplate(modelo_path) # se não der a mensagem de erro abrimos o modelo expecificado no "modelo_path"

            contexto = { #contexto dos placeholders indicados no arquivo word
                "nome": nome,
                "id": id_funcional,
                "dataDoe": data_doe,
                "dataExe": data_exe,
                "cargo": cargo,
                "genero": genero
            }

            data_hoje = datetime.now().strftime("%d-%m-%Y") #cria na data da criação uma string de data nos placehoders
            nome_arquivo = f"{nome.replace(' ', '_')}_declaracao_inss {data_hoje}.docx"
            caminho_saida = os.path.join(saida_path, nome_arquivo)

            doc.render(contexto)
            doc.save(caminho_saida)

            messagebox.showinfo("Sucesso", f"Documento gerado com sucesso:\n{caminho_saida}")
            janela.destroy()  # Fecha a janela após gerar

        except Exception as e:
            messagebox.showerror("Erro", f"Ocorreu um erro:\n{e}")

    # Interface Tkinter
    janela = tk.Tk()
    janela.title("Declaração de INSS")

    tk.Label(janela, text="Nome:").grid(row=0, column=0, sticky="e")
    entry_nome = tk.Entry(janela, width=20)
    entry_nome.grid(row=0, column=1)

    tk.Label(janela, text="ID Funcional:").grid(row=1, column=0, sticky="e")
    entry_id = tk.Entry(janela, width=20)
    entry_id.grid(row=1, column=1)

    tk.Label(janela, text="Data DOE:").grid(row=2, column=0, sticky="e")
    entry_doe = tk.Entry(janela, width=20)
    entry_doe.grid(row=2, column=1)
    entry_doe.bind("<KeyRelease>", lambda event: formatar_data(event, entry_doe))

    tk.Label(janela, text="Data Exercício:").grid(row=3, column=0, sticky="e")
    entry_exe = tk.Entry(janela, width=20)
    entry_exe.grid(row=3, column=1)
    entry_exe.bind("<KeyRelease>", lambda event: formatar_data(event, entry_exe))

    tk.Label(janela, text="Setor:").grid(row=4, column=0, sticky="e")
    entry_cargo = tk.Entry(janela, width=20)
    entry_cargo.grid(row=4, column=1)

    # genero com radiusButtons
    genero_var = tk.StringVar(value="Masculino") # valor padrão
    tk.Label(janela, text="Genero:").grid(row=6, column=0, sticky="e")
    tk.Radiobutton(janela, text="Masculino", variable=genero_var, value="Masculino").grid(row=6, column=1, sticky="w")
    tk.Radiobutton(janela, text="Feminino", variable=genero_var, value="Feminino").grid(row=7, column=1, sticky="w")

    btn_gerar = tk.Button(janela, text="Gerar Documento", command=gerar_documento)
    btn_gerar.grid(row=8, column=0, columnspan=2, pady=10)

    janela.mainloop()
