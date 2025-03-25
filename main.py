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

def criar_interface():
    # Cria uma interface gráfica simples com Tkinter(pretendo conhecer outra biblioteca para formatação de janela interativa de python)
    janela = tk.Tk()
    janela.title("Processador de Planilhas")
    janela.geometry("600x250")
    janela.configure(bg="#87CEEB")  # Azul claro
    
    botao = tk.Button(
        janela, 
        text="Selecionar Arquivo", 
        command=selecionar_arquivo,
        font=("Arial", 14, "bold"),
        bg="#4CAF50",  # Verde
        fg="white",
        padx=20,
        pady=10
    )
    botao.pack(pady=50)
    
    janela.mainloop()

# Executa a interface
try:
    if __name__ == "__main__":
        criar_interface()
except Exception as e:
    print(f"Ocorreu um erro: {e}")