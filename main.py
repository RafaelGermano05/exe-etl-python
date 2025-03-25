import pandas as pd
import tkinter as tk
from tkinter import filedialog
from pathlib import Path
from datetime import datetime


def selecionar_arquivo():
#    Abre uma janela para selecionar um arquivo CSV ou Excel.
    caminho_arquivo = filedialog.askopenfilename(
        filetypes=[("Arquivos CSV", "*.csv"), ("Arquivos Excel", "*.xlsx")]
    )
    
    if caminho_arquivo:
        processar_arquivo(caminho_arquivo)

def processar_arquivo(caminho_arquivo):
    # Processa o arquivo selecionado e aplica transformações conforme o nome do arquivo.
    nome_arquivo = Path(caminho_arquivo).stem
    
    if caminho_arquivo.endswith(".csv"):
        df = pd.read_csv(caminho_arquivo, sep=";", encoding="utf-8")
    elif caminho_arquivo.endswith(".xlsx"):
        df = pd.read_excel(caminho_arquivo)
    else:
        print("Formato não suportado!")
        return
    
    
    if "faturamento" in nome_arquivo:
        df["Faturamento_M0"] = df["Faturamento_M0"].astype(str)
        df["Faturamento_M1"] = df["Faturamento_M1"].astype(str)
        df = df[df["Status"] != "Inativo"]
        df.loc[df["Nome"] == "Empresa A", "Nome"] = "Empresa X"
    
    elif "faturamento1" in nome_arquivo:
        df["Lucro"] = df["Receita_Mensal"] - df["Despesa_Mensal"]
        df = df[df["Situação"] == "Ativo"]
        df.rename(columns={"Empresa": "Nome_Empresa"}, inplace=True)
    else:
        print("Nenhuma transformação específica para esse arquivo.")
    
    #Usei o timestamp para renomear o arquivo de acordo com a data salva
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    caminho_saida = Path(caminho_arquivo).parent / f"{nome_arquivo}_tratado_{timestamp}.xlsx"
    
    
    df.to_excel(caminho_saida, index=False)
    print(f"Arquivo salvo em: {caminho_saida}")

# Usei Tkinter mas pretendo conhecer outras bibliotecas para estilizar interface em python, sabe é o boss
janela = tk.Tk()
janela.title("Processador de Planilhas")
janela.geometry("700x300") 
janela.configure(bg="#00099b") 

# funções para hover
def on_enter(e):
    botao_selecionar.config(bg="#487d5c")  

def on_leave(e):
    botao_selecionar.config(bg="#4CAF50")  

botao_selecionar = tk.Button(
    janela, 
    text="Selecionar Planilha", 
    command=selecionar_arquivo,
    font=("Verdana", 17, "bold"),
    bg="#4CAF50",  
    fg="white",  
    padx=55,
    pady=30,
    cursor="hand2"  
)


botao_selecionar.bind("<Enter>", on_enter)
botao_selecionar.bind("<Leave>", on_leave)

botao_selecionar.pack(pady=50)  

janela.mainloop()