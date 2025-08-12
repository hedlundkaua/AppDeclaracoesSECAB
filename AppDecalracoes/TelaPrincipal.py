import tkinter as tk
from tkinter import messagebox
from datetime import datetime
from telas import Tela_INSS, Tela_Vinculo, Tela_VinculoCC, Tela_NegativaPADServidorAtivo, Tela_NegativaPADServidorInativo


# dicionario de modelos
modelos = {
    "INSS": Tela_INSS.abrir_tela,
    "Vinculo": Tela_Vinculo.abrir_tela,
    "Vinculo CC": Tela_VinculoCC.abrir_tela,
    "Negativa PAD (Servidor Ativo)": Tela_NegativaPADServidorAtivo.abrir_tela,
    "Negativa PAD (Servidor Inativo)": Tela_NegativaPADServidorInativo.abrir_tela
}

def abrir_modelo(modelo):
    janela.destroy()
    if modelo in modelos:
        modelos[modelo]()
    else: 
        messagebox.showerror("Erro", "Modelo não encontrado.")

janela = tk.Tk()
janela.title("Gerador d declarações")
janela.geometry("350x400")

tk.Label(janela, text="Escolha o tipo de declaração: ", font=("Arial", 12)).pack(pady=20)

for modelo in modelos:
    tk.Button(janela, text=modelo, width=25, height=2, command= lambda m = modelo: abrir_modelo(m)).pack(pady=5)

janela.mainloop()
    