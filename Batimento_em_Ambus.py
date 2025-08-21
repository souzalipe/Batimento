import pandas as pd
from pathlib import Path
import re
import tkinter as tk
from tkinter import filedialog, messagebox

def normaliza_cnpj(cnpj):
    cnpj_limpo = re.sub(r"\D", "", str(cnpj))
    return cnpj_limpo if len(cnpj_limpo) == 14 else None

def formatar_cnpj(cnpj):
    cnpj = normaliza_cnpj(cnpj)
    if cnpj:
        return f"{cnpj[:2]}.{cnpj[2:5]}.{cnpj[5:8]}/{cnpj[8:12]}-{cnpj[12:]}"
    return None

def carregar_excel(path):
    try:
        return pd.read_excel(path, engine="openpyxl", dtype=str)
    except Exception as e:
        messagebox.showerror("Erro", f"Erro ao carregar o arquivo {path}:\n{e}")
        return None

def filtrar_cadfi(df):
    required = ["Administrador", "Situacao", "Tipo_Fundo", "Denominacao_Social", "CNPJ_Fundo"]
    if not all(col in df.columns for col in required):
        messagebox.showerror("Erro", f"Colunas ausentes no CadFi: {set(required) - set(df.columns)}")
        return None
    df = df[
        (df["Administrador"].fillna("") == "BB GESTAO DE RECURSOS DTVM S.A") &
        (df["Situacao"] == "Em Funcionamento Normal") &
        (df["Tipo_Fundo"] == (["FI", "FAPI", "FMP-FGTS", "FIIM","Findice"])) &
        (~df["Denominacao_Social"].str.contains(
            "fic|cotas|FIC de FI|fic de fi|fi de fic|FC|fc|"
            "BB FUNDO DE INVESTIMENTO RENDA FIXA DAS PROVISÕES TÉCNICAS DOS CONSÓRCIOS DO SEGURO DPVAT|"
            "BB RJ FUNDO DE INVESTIMENTO MULTIMERCADO|"
            "BB ZEUS MULTIMERCADO FUNDO DE INVESTIMENTO|"
            "BB AQUILES FUNDO DE INVESTIMENTO RENDA FIXA|"
            "BRASILPREV FIX ESTRATÉGIA 2025 III FIF FIF RENDA FIXA RESPONSABILIDADE LIMITADA",
            case=False, na=False))
    ].copy()
    df["CNPJ"] = df["CNPJ_Fundo"].apply(formatar_cnpj)
    df = df[df["CNPJ"].notnull()]
    return df.drop_duplicates(subset="CNPJ")

def carregar_controle(path):
    df = carregar_excel(path)
    if df is not None and "CNPJ" in df.columns:
        df["CNPJ"] = df["CNPJ"].apply(normaliza_cnpj).apply(formatar_cnpj)
        df = df[df["CNPJ"].notnull()]
        return df.drop_duplicates(subset="CNPJ")
    messagebox.showerror("Erro", "Coluna 'CNPJ' ausente no Controle Espelho.")
    return None

def comparar_fundos(cadfi_df, controle_df):
    if cadfi_df is None or controle_df is None:
        return None
    return cadfi_df[cadfi_df["CNPJ"].isin(set(controle_df["CNPJ"]))].copy()

def gerar_relatorio(df, path):
    if df is None or df.empty:
        messagebox.showinfo("Resultado", "Nenhum fundo encontrado em ambos os arquivos.")
        return
    df["GFI"] = df.get("GFI", "Não possui")
    relatorio_df = df[["CNPJ", "Denominacao_Social", "GFI"]].rename(columns={
        "Denominacao_Social": "Nome do fundo",
        "GFI": "Número de Protocolo (GFI)"
    })
    try:
        relatorio_df.to_excel(path, index=False, engine="openpyxl")
        messagebox.showinfo("Sucesso", f"Relatório salvo em:\n{path.resolve()}")
    except Exception as e:
        messagebox.showerror("Erro", f"Erro ao salvar o relatório:\n{e}")

def selecionar_arquivo(entry):
    file = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])
    if file:
        entry.delete(0, tk.END)
        entry.insert(0, file)

def iniciar_processo():
    cadfi_path = Path(entry_cadfi.get())
    controle_path = Path(entry_controle.get())
    relatorio_path = Path("Fundos_em_ambos.xlsx")

    cadfi_df = carregar_excel(cadfi_path)
    cadfi_filtrado = filtrar_cadfi(cadfi_df)
    controle_df = carregar_controle(controle_path)

    fundos_em_ambos = comparar_fundos(cadfi_filtrado, controle_df)
    gerar_relatorio(fundos_em_ambos, relatorio_path)

# Interface gráfica
root = tk.Tk()
root.title("Batimento de Fundos - Em Ambos")
root.geometry("650x220")
root.resizable(False, False)

# Centraliza a janela
root.update_idletasks()
width = root.winfo_width()
height = root.winfo_height()
x = (root.winfo_screenwidth() // 2) - (width // 2)
y = (root.winfo_screenheight() // 2) - (height // 2)
root.geometry(f"{width}x{height}+{x}+{y}")

# Título
titulo = tk.Label(root, text="Batimento de Fundos - Presentes em Ambos", font=("Helvetica", 14, "bold"))
titulo.pack(pady=10)

# Frame principal
frame = tk.Frame(root)
frame.pack()

#w", width=22).grid(row=0, column=0, padx=5, pady=5)
entry_cadfi = tk.Entry(frame, width=50)
entry_cadfi.grid(row=0, column=1, padx=5)
tk.Button(frame, text="Selecionar", command=lambda: selecionar_arquivo(entry_cadfi)).grid(row=0, column=2, padx=5)

# Linha 2 - Controle Espelho
tk.Label(frame, text="Arquivo Controle Espelho:", anchor="w", width=22).grid(row=1, column=0, padx=5, pady=5)
entry_controle = tk.Entry(frame, width=50)
entry_controle.grid(row=1, column=1, padx=5)
tk.Button(frame, text="Selecionar", command=lambda: selecionar_arquivo(entry_controle)).grid(row=1, column=2, padx=5)

# Botão de iniciar
btn_iniciar = tk.Button(root, text="Gerar Relatório", command=iniciar_processo, bg="#1BA8E0", fg="black", font=("Helvetica", 10, "bold"))
btn_iniciar.pack(pady=15)

root.mainloop()
