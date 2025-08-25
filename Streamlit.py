# app.py
import io
import re
import unicodedata
from pathlib import Path
import pandas as pd
import streamlit as st

# ========= FUNÇÕES DE SUPORTE (baseadas no seu script atual) =========
def normaliza_cnpj(cnpj):
    cnpj_limpo = re.sub(r"\D", "", str(cnpj))
    return cnpj_limpo if len(cnpj_limpo) == 14 else None

def formatar_cnpj(cnpj):
    cnpj = normaliza_cnpj(cnpj)
    if cnpj:
        return f"{cnpj[:2]}.{cnpj[2:5]}.{cnpj[5:8]}/{cnpj[8:12]}-{cnpj[12:]}"
    return None

def remover_duplicatas_por_cnpj(df, coluna_origem):
    df = df.copy()
    df["CNPJ_Normalizado"] = df[coluna_origem].apply(normaliza_cnpj)
    df["CNPJ"] = df["CNPJ_Normalizado"].apply(formatar_cnpj)
    df = df[df["CNPJ"].notnull()]
    return df.drop_duplicates(subset="CNPJ").copy()

def padronizar_colunas(df):
    df = df.copy()
    def norm(s):
        s = unicodedata.normalize("NFKD", str(s))
        s = s.encode("ascii", "ignore").decode("utf-8")
        s = s.strip()
        return s
    df.columns = [norm(c) for c in df.columns]
    return df

def carregar_excel(arquivo):
    # 'arquivo' pode ser um UploadedFile (Streamlit) ou caminho
    df = pd.read_excel(arquivo, engine="openpyxl", dtype=str)
    return padronizar_colunas(df)

def filtrar_cadfi(df):
    required = ["Administrador", "Situacao", "Tipo_Fundo", "Denominacao_Social", "CNPJ_Fundo"]
    if not all(col in df.columns for col in required):
        faltantes = set(required) - set(df.columns)
        raise ValueError(f"Colunas ausentes no CadFi: {faltantes}")

    filtro = (
        (df["Administrador"].fillna("") == "BB GESTAO DE RECURSOS DTVM S.A") &
        (df["Situacao"] == "Em Funcionamento Normal") &
        (df["Tipo_Fundo"].isin(["FI", "FAPI", "FIIM"])) &  # FIIM = ETF
        (~df["Denominacao_Social"].str.contains(
            "fic|cotas|FIC de FI|fic de fi|fi de fic|FC|fc|"
            "BB FUNDO DE INVESTIMENTO RENDA FIXA DAS PROVISÕES TÉCNICAS DOS CONSÓRCIOS DO SEGURO DPVAT|"
            "BB RJ FUNDO DE INVESTIMENTO MULTIMERCADO|"
            "BB ZEUS MULTIMERCADO FUNDO DE INVESTIMENTO|"
            "BB AQUILES FUNDO DE INVESTIMENTO RENDA FIXA|"
            "BRASILPREV FIX ESTRATÉGIA 2025 III FIF FIF RENDA FIXA RESPONSABILIDADE LIMITADA",
            case=False, na=False
        ))
    )
    df_filtrado = df.loc[filtro].copy()
    return remover_duplicatas_por_cnpj(df_filtrado, "CNPJ_Fundo")

def carregar_controle(df_controle):
    if "CNPJ" not in df_controle.columns:
        raise ValueError("Coluna 'CNPJ' ausente no Controle Espelho.")
    return remover_duplicatas_por_cnpj(df_controle, "CNPJ")

def comparar_cnpjs(cadfi_df, controle_df):
    """CadFi que NÃO estão no Controle"""
    return cadfi_df[~cadfi_df["CNPJ"].isin(set(controle_df["CNPJ"]))].copy()

def comparar_fundos_em_comum(cadfi_df, controle_df):
    """CadFi que também estão no Controle (somente interseção por CNPJ)"""
    return cadfi_df[cadfi_df["CNPJ"].isin(set(controle_df["CNPJ"]))].copy()

def relatorio_fora_controle(df):
    """Retorna DF pronto para exportar (CadFi fora do Controle)"""
    if df is None or df.empty:
        return pd.DataFrame(columns=["CNPJ", "Nome do fundo", "Número de Protocolo (GFI)"])
    df = df.copy()
    df["GFI"] = df.get("GFI", "Não possui")
    rel = df[["CNPJ", "Denominacao_Social", "GFI"]].rename(columns={
        "Denominacao_Social": "Nome do fundo",
        "GFI": "Número de Protocolo (GFI)"
    })
    return rel

def relatorio_em_comum(df):
    """
    Retorna DF pronto para exportar (CadFi em comum com Controle)
    Inclui duas colunas vazias: 'Protocolo' e 'Mes de Referencia'.
    """
    if df is None or df.empty:
        return pd.DataFrame(columns=["CNPJ", "Nome do fundo", "Número de Protocolo (GFI)", "Protocolo", "Mes de Referencia"])
    df = df.copy()
    if "GFI" not in df.columns:
        df["GFI"] = "Não possui"
    # Colunas vazias conforme solicitado
    df["Protocolo"] = ""
    df["Mes de Referencia"] = ""
    rel = df[["CNPJ", "Denominacao_Social", "GFI", "Protocolo", "Mes de Referencia"]].rename(columns={
        "Denominacao_Social": "Nome do fundo",
        "GFI": "Número de Protocolo (GFI)"
    })
    return rel

def to_excel_bytes(df, sheet_name="Relatorio"):
    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name=sheet_name)
    buffer.seek(0)
    return buffer

# ========================== INTERFACE STREAMLIT ==========================
st.set_page_config(page_title="Batimento de Fundos - CadFi x Controle", page_icon="📊", layout="centered")

st.title("📊 Batimento de Fundos — CadFi x Controle")
st.caption("Interface web dos Batimentos. Faça o upload dos dois arquivos e clique em **Processar**.")

col1, col2 = st.columns(2)
with col1:
    cadfi_file = st.file_uploader("Arquivo CadFi (.xlsx)", type=["xlsx"], accept_multiple_files=False)
with col2:
    controle_file = st.file_uploader("Arquivo Controle Espelho (.xlsx)", type=["xlsx"], accept_multiple_files=False)

processar = st.button("Processar", type="primary")

if processar:
    if not cadfi_file or not controle_file:
        st.error("⚠️ Envie os dois arquivos (CadFi e Controle Espelho) antes de processar.")
        st.stop()

    try:
        # Carrega e prepara
        cadfi_raw = carregar_excel(cadfi_file)
        controle_raw = carregar_excel(controle_file)

        cadfi_filtrado = filtrar_cadfi(cadfi_raw)
        controle_prep = carregar_controle(controle_raw)

        # Comparações
        df_fora = comparar_cnpjs(cadfi_filtrado, controle_prep)
        df_comum = comparar_fundos_em_comum(cadfi_filtrado, controle_prep)

        # Relatórios prontos
        rel_fora = relatorio_fora_controle(df_fora)
        rel_comum = relatorio_em_comum(df_comum)

        # Exibe contagens
        st.success(f"✅ Em comum: {len(rel_comum)} fundo(s)")
        st.warning(f"❌ Fora do Controle: {len(rel_fora)} fundo(s)")

        # Prévia das tabelas
        with st.expander("Visualizar — Fundos em Comum"):
            st.dataframe(rel_comum, use_container_width=True, hide_index=True)
        with st.expander("Visualizar — Fundos fora do Controle"):
            st.dataframe(rel_fora, use_container_width=True, hide_index=True)

        # Downloads
        c1, c2 = st.columns(2)
        with c1:
            st.download_button(
                label="⬇️ Baixar Relatório — Em Comum (Excel)",
                data=to_excel_bytes(rel_comum, sheet_name="Em_Comum"),
                file_name="Relatorio_Fundos_Em_Comum.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )
        with c2:
            st.download_button(
                label="⬇️ Baixar Relatório — Fora do Controle (Excel)",
                data=to_excel_bytes(rel_fora, sheet_name="Fora_do_Controle"),
                file_name="Relatorio_Fundos_Fora_do_Controle.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )

    except Exception as e:
        st.exception(e)
        st.stop()

st.markdown("---")
st.caption("Dica: se as colunas do Excel vierem com acentos/variações, o app normaliza nomes para evitar erros.")
