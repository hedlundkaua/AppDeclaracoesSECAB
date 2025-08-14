import tkinter as tk
def formatar_data(event, entry): 
        texto = entry.get().replace("/", "")[:8] #[:8] limita o texto a 8 caracteres
        #entry.get() pega o texto atual do campo, replace("/", "") remove as barras que o usuario ja digitou para evitar duplicar barras
        formatado = ""

        # para percorrer cada caractere do texto
        for i, c in enumerate(texto): 
            if i ==2 or i == 4: #quando o i 2 ou 4 (depois do dia e do mes) ele adiciona a barra 
                formatado += "/"
            formatado += c
        entry.delete(0, tk.END)
        entry.insert(0, formatado)


def formatar_cpf(event, entry):
     texto = entry.get().replace(".", "").replace("-", "")[:11]
     formato = ""
     for i, c in enumerate(texto):
        if i == 3 or i == 6:
             formato += "."
        if i == 9:
            formato += "-"
        formato += c
        entry.delete(0, tk.END)
        entry.insert(0, formato)     