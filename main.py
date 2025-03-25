from pathlib import Path
from datetime import datetime
import pandas as pd
import tkinter as tk
from tkinter import filedialog

def selecionar_arquivo():
    arquivo = filedialog.askopenfilename(filetypes=[("CSV files", "*.csv"), ("Excel files", "*.xlsx")])
    if arquivo:
        processar_planilha(arquivo)

def processar_planilha(arquivo):
    nome_arquivo = Path(arquivo).stem 
    
    # detectar formato do arquivo
    if arquivo.endswith(".csv"):
        df = pd.read_csv(arquivo, sep=";", encoding="utf-8")
    elif arquivo.endswith(".xlsx"):
        df = pd.read_excel(arquivo)
    else:
        print("Formato de arquivo não suportado")
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
        print("Nome de arquivo não reconhecido, aplicando tratamento padrão.")
    
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    caminho_saida = Path(arquivo).parent / f"{nome_arquivo}_tratado_{timestamp}.xlsx"
    
    df.to_excel(caminho_saida, index=False)
    print(f"Arquivo processado e salvo como {caminho_saida}")

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