import tkinter as tk
from tkinter import messagebox
from datetime import datetime
from telas import Tela_INSS, Tela_Vinculo, Tela_VinculoCC, Tela_NegativaPADServidorAtivo, Tela_NegativaPADServidorInativo, Tela_IsencaoDeTaxaRenovacaoCNH, Tela_PosseEFetivoExercicio, InscricaoOABeDeclaracaoDePoderDdecisorio, Tela_DasCompetenciasDosAE, Tela_DasCompetenciasDosAFRE
from telas import Tela_DasCompetenciasDosATREs

# dicionario de modelos
modelos = {
    "INSS": Tela_INSS.abrir_tela,
    "Vinculo": Tela_Vinculo.abrir_tela,
    "Vinculo CC": Tela_VinculoCC.abrir_tela,
    "Negativa PAD (Servidor Ativo)": Tela_NegativaPADServidorAtivo.abrir_tela,
    "Negativa PAD (Servidor Inativo)": Tela_NegativaPADServidorInativo.abrir_tela,
    "Isenção de CNH": Tela_IsencaoDeTaxaRenovacaoCNH.abrir_tela,
    "Posse e Efetivo Exercicio": Tela_PosseEFetivoExercicio.abrir_tela,
    "Inscrição OAB e Declaração de Poder Decisorio":InscricaoOABeDeclaracaoDePoderDdecisorio.abrir_tela,
    "Competencias dos AEs": Tela_DasCompetenciasDosAE.abrir_tela,
    "Competencias dos AFREs": Tela_DasCompetenciasDosAFRE.abrir_tela,
    "Competencias dos ATREs": Tela_DasCompetenciasDosATREs.abrir_tela
}

def abrir_modelo(modelo):
    janela.destroy()
    if modelo in modelos:
        modelos[modelo]()
    else: 
        messagebox.showerror("Erro", "Modelo não encontrado.")

janela = tk.Tk()
janela.title("Gerador d declarações")
janela.geometry("400x800")

tk.Label(janela, text="Escolha o tipo de declaração: ", font=("Arial", 12)).pack(pady=20)

for modelo in modelos:
    tk.Button(janela, text=modelo, width=35, height=2, command= lambda m = modelo: abrir_modelo(m)).pack(pady=5)

janela.mainloop()
    