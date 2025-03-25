import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox
from pathlib import Path
from datetime import datetime
from xlsxwriter import Workbook

def selecionar_arquivo():
    # Abre uma janela para selecionar um arquivo CSV ou Excel.
    caminho_arquivo = filedialog.askopenfilename(
        filetypes=[("Arquivos CSV", "*.csv"), ("Arquivos Excel", "*.xlsx")]
    )
    
    if caminho_arquivo:
        processar_arquivo(caminho_arquivo)

def processar_arquivo(caminho_arquivo):
    # Criar pasta para armazenar arquivos tratados
    pasta_saida = Path.home() / "Documentos" / "Planilhas_Tratadas"
    pasta_saida.mkdir(parents=True, exist_ok=True)
    
    nome_arquivo = Path(caminho_arquivo).stem
    
    if caminho_arquivo.endswith(".csv"):
        df = pd.read_csv(caminho_arquivo, sep=",", encoding="utf-8")  # Alterado para separação por vírgulas
    elif caminho_arquivo.endswith(".xlsx"):
        df = pd.read_excel(caminho_arquivo)
    else:
        messagebox.showerror("Erro", "Formato de arquivo não suportado!")
        return
    
    # Aplicando transformações conforme o nome do arquivo
    if "PARCEIRO ESTRATÉGICO_Principal_Tabla" in nome_arquivo:
        for col in ['TPV M0', 'TPV M1', 'TPV M2']:
            if col in df.columns:
                df[col] = df[col].astype("str")
        df.loc[df['SELLER'] == 'MODA 20', 'EXECUTIVO'] = 'PAULO RICARDO'

    elif "PARCEIRO ESTRATÉGICO_2025 - Analítico" in nome_arquivo:
        for col in ['TPV M0', 'TPV M1', 'TPV M2', 'TPV_PARCELADO_M0', 'TPV_PARCELADO_M1' ,'TPV_PARCELADO_M2', 'TPV_CREDITO_M0',
                    'TPV_CREDITO_M1', 'TPV_CREDITO_M2', 'TPV_DEBITO_M0', 'TPV_DEBITO_M1', 'TPV_DEBITO_M2']:
            if col in df.columns:
                df[col] = df[col].astype("str")

    elif "Visitas" in nome_arquivo:
        df["Data"] = df["Data"].str.replace(" feb ","/02/")
        df["Data"] = df["Data"].str.replace(" oct ","/10/")
        df["Data"] = df["Data"].str.replace(" dic ","/12/")
        df["Data"] = df["Data"].str.replace(" ene ","/01/")
        df["Data"] = df["Data"].str.replace(" mar ","/03/")
        df["Data"] = df["Data"].str.replace(" sept ","/09/")
    else:
        messagebox.showinfo("Aviso", "Nenhuma transformação específica para esse arquivo.")
    
    # usando timestamp pra nomear por data e horario de geração 
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    caminho_saida = pasta_saida / f"{nome_arquivo}_tratado_{timestamp}.xlsx"
    
    with pd.ExcelWriter(caminho_saida, engine="xlsxwriter") as writer:
        df.to_excel(writer, sheet_name="Dados Processados", index=False)
    
    messagebox.showinfo("Sucesso", f"Arquivo salvo em:\n{caminho_saida}")

# Criando a interface gráfica com Tkinter pretendo conhecer outros frames para interface
janela = tk.Tk()
janela.title("Processador de Planilhas")
janela.geometry("700x300") 
janela.configure(bg="#00099b") 

# Funções para hover

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
