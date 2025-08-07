import tkinter as tk
from tkinter import messagebox
from docxtpl import DocxTemplate
import os

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
        nome = entry_nome.get()
        id_funcional = entry_id.get()
        data_doe = entry_doe.get()
        data_exe = entry_exe.get()
        cargo = entry_cargo.get()

        if not nome or not id_funcional or not data_doe or not data_exe or not cargo:
            messagebox.showwarning("Campos obrigatórios", "Por favor, preencha todos os campos.")
            return

        # Caminho do modelo
        modelo_path = r"W:\DRH\SECAB\Kaua Teste\inss.docx"

        if not os.path.exists(modelo_path):
            messagebox.showerror("Erro", f"Modelo não encontrado:\n{modelo_path}")
            return

        doc = DocxTemplate(modelo_path)

        contexto = {
            "nome": nome,
            "id": id_funcional,
            "dataDoe": data_doe,
            "dataExe": data_exe,
            "cargo": cargo
        }

        pasta_saida = os.path.dirname(modelo_path)
        nome_arquivo = f"{nome.replace(' ', '_')}_declaracao.docx"
        caminho_saida = os.path.join(pasta_saida, nome_arquivo)

        doc.render(contexto)
        doc.save(caminho_saida)

        messagebox.showinfo("Sucesso", f"Documento gerado com sucesso:\n{caminho_saida}")
        janela.destroy()  # Fecha a janela após gerar

    except Exception as e:
        messagebox.showerror("Erro", f"Ocorreu um erro:\n{e}")

# Interface Tkinter
janela = tk.Tk()
janela.title("Gerador de Declaração")

tk.Label(janela, text="Nome:").grid(row=0, column=0, sticky="e")
entry_nome = tk.Entry(janela, width=40)
entry_nome.grid(row=0, column=1)

tk.Label(janela, text="ID Funcional:").grid(row=1, column=0, sticky="e")
entry_id = tk.Entry(janela, width=40)
entry_id.grid(row=1, column=1)

tk.Label(janela, text="Data DOE (dd/mm/aaaa):").grid(row=2, column=0, sticky="e")
entry_doe = tk.Entry(janela, width=40)
entry_doe.grid(row=2, column=1)
entry_doe.bind("<KeyRelease>", lambda event: formatar_data(event, entry_doe))

tk.Label(janela, text="Data Exercício (dd/mm/aaaa):").grid(row=3, column=0, sticky="e")
entry_exe = tk.Entry(janela, width=40)
entry_exe.grid(row=3, column=1)
entry_exe.bind("<KeyRelease>", lambda event: formatar_data(event, entry_exe))

tk.Label(janela, text="Cargo:").grid(row=4, column=0, sticky="e")
entry_cargo = tk.Entry(janela, width=40)
entry_cargo.grid(row=4, column=1)

btn_gerar = tk.Button(janela, text="Gerar Documento", command=gerar_documento)
btn_gerar.grid(row=5, column=0, columnspan=2, pady=10)

janela.mainloop()
