import io
import re
import unicodedata
from pathlib import Path
import pandas as pd
import streamlit as st



def so_digitos(s):
    return re.sub(r'\D', '', str(s or ''))

def normaliza_cnpj(cnpj):
    d = so_digitos(cnpj)
    if len(d) == 14:
        return d
    # se vier com menos dígitos, completa com zeros à esquerda
    if 0 < len(d) < 14:
        return d.zfill(14)
    return None

def formatar_cnpj(cnpj):
    d = normaliza_cnpj(cnpj)
    if not d or len(d) != 14:
        return None
    return f"{d[:2]}.{d[2:5]}.{d[5:8]}/{d[8:12]}-{d[12:]}"


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


def normaliza_texto(s):
    s = str(s or "")
    s = unicodedata.normalize("NFKD", s).encode("ascii", "ignore").decode("utf-8")
    return s.strip().upper()

def _norm_header_key(s: str) -> str:
    s = unicodedata.normalize("NFKD", str(s)).encode("ascii", "ignore").decode("utf-8")
    s = re.sub(r"\s+", " ", s.strip().lower())
    s = re.sub(r"[^a-z0-9]+", "_", s)
    s = re.sub(r"_+", "_", s).strip("_")
    return s

def _encontrar_coluna_status(df: pd.DataFrame):
    """
    Tenta localizar a coluna que representa a 'situação/status' do fundo
    (ex.: 'Situação', 'Situacao', 'Status', 'Status do Fundo', etc.)
    Retorna o nome ORIGINAL da coluna ou None se não encontrar.
    """
    norm_map = {_norm_header_key(c): c for c in df.columns}

    # 1) Prioridades mais comuns
    prioridade = [
        "situacao", "situação", "situacao_do_fundo", "situcao", "status", "status_do_fundo"
    ]
    for key in prioridade:
        if key in norm_map:
            return norm_map[key]

    # 2) Heurística por palavras-chave
    candidatos = []
    for k, original in norm_map.items():
        score = 0
        if "situac" in k or "situa" in k: score += 3
        if "status" in k:                 score += 2
        if "fundo" in k:                  score += 1
        if "cnpj" in k:                   score = -1
        if score > 0:
            candidatos.append((score, original))
    if candidatos:
        candidatos.sort(reverse=True, key=lambda x: x[0])
        return candidatos[0][1]

    return None

# Conjunto de valores considerados "ativos" após normalização (UPPER + sem acentos)
VALORES_ATIVOS = {
    normaliza_texto("Em Funcionamento Normal"),
    normaliza_texto("Em Funcionamento"),
    normaliza_texto("Ativo"),
    normaliza_texto("Ativa"),
    normaliza_texto("Em Atividade"),
    normaliza_texto("A"),
}

def filtrar_status_ativos(df: pd.DataFrame) -> pd.DataFrame:
    """
    Mantém apenas linhas cuja situação/status esteja na lista 'VALORES_ATIVOS'.
    Se não encontrar a coluna, retorna o df sem alterações (fail-safe).
    """
    if df is None or df.empty:
        return df

    col = _encontrar_coluna_status(df)
    if not col:
        # Coluna de status não encontrada; não filtra para evitar perda de dados
        return df

    out = df.copy()
    out["_STATUS_NORM_"] = out[col].map(normaliza_texto)
    out = out[out["_STATUS_NORM_"].isin(VALORES_ATIVOS)].drop(columns=["_STATUS_NORM_"])
    return out


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
            "BRASILPREV FIX ESTRATÉGIA 2025 III FIF FIF RENDA FIXA RESPONSABILIDADE LIMITADA"
            "BB MASTER RENDA FIXA DEBÊNTURES INCENTIVADAS FIF INVESTIMENTO EM INFRAESTRUTURA RESP LIMITADA"
            "BB CIN"
            "BB BNC AÇÕES NOSSA CAIXA NOSSO CLUBE DE INVESTIMENTO",
            case=False, na=False
        ))
    )
    df_filtrado = df.loc[filtro].copy()
    return remover_duplicatas_por_cnpj(df_filtrado, "CNPJ_Fundo")

def comparar_controle_fora_cadfi(cadfi_df, controle_df):
    """
    Retorna registros do Controle cujo CNPJ não aparece no CadFi (após padronização/duplicatas).
    Espera que ambas as tabelas já tenham a coluna 'CNPJ' formatada.
    """
    return controle_df[~controle_df["CNPJ"].isin(set(cadfi_df["CNPJ"]))].copy()



def _norm_header_key(s: str) -> str:
    """
    Normaliza rótulos de coluna: remove acentos, baixa, troca não-alfanum por '_'.
    Ex.: 'Denominação Social' -> 'denominacao_social'
    """
    s = unicodedata.normalize("NFKD", str(s)).encode("ascii", "ignore").decode("utf-8")
    s = re.sub(r"\s+", " ", s.strip().lower())
    s = re.sub(r"[^a-z0-9]+", "_", s)
    s = re.sub(r"_+", "_", s).strip("_")
    return s

def _encontrar_coluna_nome(df: pd.DataFrame) -> str:
    """
    Tenta localizar a coluna que representa o 'nome do fundo' no Controle,
    usando uma lista de prioridades + heurística por palavras-chave.
    Retorna o nome ORIGINAL da coluna (não normalizado), ou None.
    """
    # Mapa: chave_normalizada -> coluna_original
    norm_map = {_norm_header_key(c): c for c in df.columns}

    # 1) Prioridades mais comuns
    prioridade = [
        "denominacao_social", "denominacao_do_fundo", "denominacao",
        "nome_do_fundo", "nome_fundo", "nome",
        "razao_social", "razao", "descricao"
    ]
    for key in prioridade:
        if key in norm_map:
            return norm_map[key]

    # 2) Heurística: qualquer coluna com 'denomin' ou 'nome' e (idealmente) 'fundo'
    candidatos = []
    for k, original in norm_map.items():
        score = 0
        if "denomin" in k: score += 3
        if "nome" in k:    score += 2
        if "fundo" in k:   score += 1
        if "cnpj" in k:    score = -1  # nunca usar CNPJ como nome
        if score > 0:
            candidatos.append((score, original))
    if candidatos:
        candidatos.sort(reverse=True, key=lambda x: x[0])
        return candidatos[0][1]

    # 3) Fallback: primeira coluna de texto que não seja 'CNPJ'
    for c in df.columns:
        if c != "CNPJ" and df[c].dtype == object:
            return c
    return None

def relatorio_controle_fora_cadfi(df_controle: pd.DataFrame) -> pd.DataFrame:
    """
    Monta DataFrame pronto para exportar com base no Controle que NÃO está no CadFi.
    Inclui 'CNPJ' e 'Nome do fundo (Controle)', detectando a melhor coluna de nome.
    """
    if df_controle is None or df_controle.empty:
        return pd.DataFrame(columns=["CNPJ", "Nome do fundo (Controle)"])

    out = df_controle.copy()
    col_nome = _encontrar_coluna_nome(out)

    if col_nome and col_nome in out.columns:
        out = out.rename(columns={col_nome: "Nome do fundo (Controle)"})
        out["Nome do fundo (Controle)"] = (
            out["Nome do fundo (Controle)"]
            .astype(str)
            .str.strip()
        )
    else:
        # não encontrou, exporta coluna vazia (mas mantém estrutura)
        out["Nome do fundo (Controle)"] = ""

    return out[["CNPJ", "Nome do fundo (Controle)"]]

# --- NOVO: excluir fundos indesejados pelo nome (apenas no Controle fora do CadFi) ---
EXCLUIR_NOMES_CONTROLE = [
    "BB CIN",
    "BB BNC AÇÕES NOSSA CAIXA NOSSO CLUBE DE INVESTIMENTO",
]

def filtrar_controle_por_nome(df: pd.DataFrame,
                              nomes_excluir=EXCLUIR_NOMES_CONTROLE) -> pd.DataFrame:
    """
    Remove do DF quaisquer linhas cujo 'nome do fundo' (coluna detectada)
    contenha algum dos nomes/padrões informados (case/acentos-insensitive).
    """
    if df is None or df.empty:
        return df

    # Reusa seu detector de coluna de nome
    col_nome = _encontrar_coluna_nome(df)
    if not col_nome or col_nome not in df.columns:
        # Falha silenciosa: se não achar a coluna de nome, não filtra
        return df

    nomes_norm = [normaliza_texto(n) for n in nomes_excluir]

    out = df.copy()
    out["_NOME_NORM_"] = out[col_nome].map(normaliza_texto)

    # Exclui se QUALQUER padrão aparecer no nome normalizado
    mask_excluir = out["_NOME_NORM_"].apply(lambda s: any(p in s for p in nomes_norm))
    out = out[~mask_excluir].drop(columns=["_NOME_NORM_"])

    return out

# --- NOVO: excluir fundos do Controle por SITUACAO ('I' e 'P') ---
EXCLUIR_SITUACAO_CONTROLE = ("I", "P")

def filtrar_controle_por_situacao(df: pd.DataFrame,
                                  excluir_codigos=EXCLUIR_SITUACAO_CONTROLE) -> pd.DataFrame:
    """
    Remove linhas do DF cujo 'status/situacao' seja 'I' ou 'P'.
    - Detecção da coluna via _encontrar_coluna_status.
    - Normaliza texto (sem acento/maiúsculas) e usa a 1ª letra como código.
    - Ex.: 'Inativo' -> 'I', 'Paralisado' -> 'P'.
    """
    if df is None or df.empty:
        return df

    col_status = _encontrar_coluna_status(df)
    if not col_status or col_status not in df.columns:
        # Falha silenciosa se não achar coluna de situação/status
        return df

    excluir_norm = {normaliza_texto(x)[:1] for x in excluir_codigos}

    out = df.copy()
    # Extrai a 1ª letra do status normalizado (ou string vazia)
    out["_SIT_"] = out[col_status].map(lambda x: normaliza_texto(x)[:1] if pd.notna(x) else "")
    mask_excluir = out["_SIT_"].isin(excluir_norm)

    out = out[~mask_excluir].drop(columns=["_SIT_"])
    return out


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
        return pd.DataFrame(columns=["CNPJ", "Nome do fundo"])
    df = df.copy()
    rel = df[["CNPJ", "Denominacao_Social"]].rename(columns={
        "Denominacao_Social": "Nome do fundo",
    })
    return rel

def relatorio_em_comum(df):
    """
    Retorna DF pronto para exportar (CadFi em comum com Controle)
    Inclui colunas: 'Protocolo-CDA', 'Mes_de_Referencia-CDA',
    'Protocolo-Balancete', 'Mes_de_Referencia-Balancete'.
    """
    if df is None or df.empty:
        return pd.DataFrame(columns=[
            "CNPJ", "Nome do fundo","Protocolo-CDA", "Mes_de_Referencia-CDA",
            "Protocolo-Balancete", "Mes_de_Referencia-Balancete"
        ])

    df = df.copy()

    # Garante que todas as colunas necessárias existam
    df["Protocolo-CDA"] = df.get("Protocolo-CDA", "")
    df["Mes_de_Referencia-CDA"] = df.get("Mes_de_Referencia-CDA", "")
    df["Protocolo-Balancete"] = df.get("Protocolo-Balancete", "")
    df["Mes_de_Referencia-Balancete"] = df.get("Mes_de_Referencia-Balancete", "")

    # Monta o relatório
    rel = df[[
        "CNPJ", "Denominacao_Social","Protocolo-CDA", "Mes_de_Referencia-CDA",
        "Protocolo-Balancete", "Mes_de_Referencia-Balancete"
    ]].rename(columns={
        "Denominacao_Social": "Nome do fundo",
    })

    return rel


def to_excel_bytes(df, sheet_name="Relatorio"):
    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name=sheet_name)
    buffer.seek(0)
    return buffer

# ======================= /CDA =====================================================

def _normaliza_competencia_mm_aaaa(valor: str) -> str:
    """
    Normaliza entradas como '2025-06-01 00:00:00', '01/06/2025' ou '06/2025' para '06/2025'.
    """
    s = str(valor or "").strip()
    if not s:
        return ""

    # Ex.: 2025-06-01 ou 2025/06/...
    m_iso = re.search(r'(20\d{2})-/', s)
    if m_iso:
        y, mth = m_iso.groups()
        return f"{int(mth):02d}/{y}"

    # Ex.: 01/06/2025 ou 06/01/2025 (decide o mês coerente)
    m_dmy = re.search(r'(\d{1,2})/-/-', s)
    if m_dmy:
        a, b, y = m_dmy.groups()
        a, b = int(a), int(b)
        mth = a if a <= 12 else b
        return f"{int(mth):02d}/{y}"

    # Ex.: 06/2025
    m_my = re.search(r'(0?[1-9]|1[0-2])/-', s)
    if m_my:
        mth, y = m_my.groups()
        return f"{int(mth):02d}/{y}"

    return s or ""


def parse_protocolos_cda_xlsx(arquivo_xlsx) -> pd.DataFrame:
    """
    Converte o Excel de protocolos (layout de texto em linhas) em um DF tabular:
      CNPJ_Masked | CNPJ_Num | Participante | CDA_Protocolo | CDA_Competencia | Data_Acao
    Mantém 1 linha por CNPJ (o protocolo mais RECENTE pela Data_Acao).
    """
    df_raw = pd.read_excel(arquivo_xlsx, sheet_name=0, header=None, dtype=str)

    # Achatar todas as células em lista de linhas não-vazias (ordem preservada)
    lines = []
    for _, row in df_raw.iterrows():
        for val in row.values:
            if pd.isna(val):
                continue
            txt = str(val).strip()
            if txt:
                lines.append((len(lines), txt))

    n = len(lines)
    registros = []

    for i in range(n):
        _, text = lines[i]
        low = text.lower()

        # Marcador do bloco do protocolo (várias variantes)
        if low.startswith('nº protocolo') or low.startswith('n° protocolo') or low.startswith('no protocolo') \
           or low.startswith('nº do protocolo') or low.startswith('n° do protocolo'):
            # 1) Pega o valor do Nº Protocolo (primeira linha "não-cabeçalho" abaixo)
            protocolo = None
            j = i + 1
            while j < n:
                _, t2 = lines[j]
                low2 = t2.lower()
                # pula rótulos/cabeçalhos usuais
                if re.match(
                    r'^(protocolo de confirm|status:|informe:|opera|documento:|compet|usuário|usuario|nº do recebimento|nome do arquivo|participante:|tipo do participante|data ação:|data acao:)',
                    low2
                ):
                    j += 1
                    continue
                # valor candidato
                protocolo = t2.strip()
                # limpa sufixo '.0' que pode vir do Excel
                if protocolo.endswith(".0"):
                    protocolo = protocolo[:-2]
                break
                j += 1

            # 2) Sobe até "Participante:" para capturar nome e CNPJ
            cnpj_masked, participante = None, None
            k = i
            while k >= 0:
                _, tprev = lines[k]
                lowp = tprev.lower()
                if lowp.startswith('participante'):
                    first_name = None
                    kk = k + 1
                    while kk < n:
                        _, tline = lines[kk]
                        low2 = tline.lower()
                        # limites do bloco
                        if low2.startswith('tipo do participante') or low2.startswith('data ação') or low2.startswith('data acao') or low2.startswith('nº protocolo') or low2.startswith('n° protocolo') or low2.startswith('nº do protocolo') or low2.startswith('n° do protocolo'):
                            break
                        # primeiro texto após "Participante:" é o nome
                        if first_name is None and tline:
                            first_name = tline.strip()
                        # CNPJ pode estar nessa linha ou na próxima (entre parênteses)
                        m = re.search(r'(\d{2}\.\d{3}\.\d{3}/\d{4}-\d{2})', tline)
                        if m:
                            cnpj_masked = m.group(1)
                            break
                        kk += 1
                    participante = first_name
                    break
                k -= 1

            # 3) Competência (acima)
            competencia_raw = None
            k = i
            while k >= 0:
                _, tprev = lines[k]
                lowp = tprev.lower()
                if lowp.startswith('competência') or lowp.startswith('competencia'):
                    kk = k + 1
                    while kk < n:
                        _, tval = lines[kk]
                        if tval:
                            competencia_raw = tval.strip()
                            break
                        kk += 1
                    break
                k -= 1

            # 4) Data Ação (acima)
            data_acao_raw = None
            k = i
            while k >= 0:
                _, tprev = lines[k]
                lowp = tprev.lower()
                if lowp.startswith('data ação') or lowp.startswith('data acao'):
                    kk = k + 1
                    while kk < n:
                        _, tval = lines[kk]
                        if tval:
                            data_acao_raw = tval.strip()
                            break
                        kk += 1
                    break
                k -= 1

            # Monta registro (se tiver CNPJ e Protocolo)
            if cnpj_masked and protocolo:
                cnpj_num = normaliza_cnpj(cnpj_masked)  # helper seu
                comp = _normaliza_competencia_mm_aaaa(competencia_raw)
                try:
                    data_acao = pd.to_datetime(data_acao_raw, dayfirst=True, errors='coerce') if data_acao_raw else pd.NaT
                except Exception:
                    data_acao = pd.NaT

                registros.append({
                    "CNPJ_Masked": cnpj_masked,
                    "CNPJ_Num": cnpj_num,
                    "Participante": participante,
                    "CDA_Protocolo": protocolo,
                    "CDA_Competencia": comp,
                    "Data_Acao": data_acao
                })

    df = pd.DataFrame(registros)
    if df.empty:
        return df

    # Mantém o protocolo mais recente por CNPJ (pela Data_Acao)
    df = df.sort_values(["CNPJ_Num", "Data_Acao"], ascending=[True, False]).drop_duplicates("CNPJ_Num", keep="first")
    return df


def enriquecer_em_comum_com_cda(rel_em_comum_df: pd.DataFrame, df_cda: pd.DataFrame) -> pd.DataFrame:
    """
    Recebe o DF 'Relatorio_Fundos_Em_Ambos' (com colunas 'CNPJ', 'Protocolo', 'Mes de Referencia' etc.)
    e adiciona as colunas 'CDA_Protocolo' e 'CDA_Competencia' posicionadas depois de 'Mes de Referencia'.
    Onde não houver match, preenche com 'Não possui'.
    """
    if rel_em_comum_df is None or rel_em_comum_df.empty:
        out = rel_em_comum_df.copy()
        if "CDA_Protocolo" not in out.columns:
            out["CDA_Protocolo"] = ""
        if "CDA_Competencia" not in out.columns:
            out["CDA_Competencia"] = ""
        return out

    if df_cda is None or df_cda.empty:
        out = rel_em_comum_df.copy()
        out["CDA_Protocolo"] = "Não possui"
        out["CDA_Competencia"] = "Não possui"
        return out

    rel = rel_em_comum_df.copy()
    rel["CNPJ_Num"] = rel["CNPJ"].map(normaliza_cnpj)

    enx = rel.merge(
        df_cda[["CNPJ_Num", "CDA_Protocolo", "CDA_Competencia"]],
        on="CNPJ_Num", how="left"
    )
    enx["CDA_Protocolo"] = enx["CDA_Protocolo"].fillna("Não possui")
    enx["CDA_Competencia"] = enx["CDA_Competencia"].fillna("Não possui")

    # Posiciona após "Mes de Referencia" (se existir), senão ao final
    cols = list(enx.columns)
    insert_pos = cols.index("Mes de Referencia") + 1 if "Mes de Referencia" in cols else len(cols)
    for newcol in ["CDA_Protocolo", "CDA_Competencia"]:
        if newcol in cols:
            cols.remove(newcol)
    cols = cols[:insert_pos] + ["CDA_Protocolo", "CDA_Competencia"] + cols[insert_pos:]
    enx = enx[cols]

    # Remove auxiliar
    if "CNPJ_Num" in enx.columns:
        enx = enx.drop(columns=["CNPJ_Num"])

    return enx

# ======================= /CDA =====================================================


# ========================== INTERFACE STREAMLIT ==========================
st.set_page_config(page_title="Batimento de Fundos - CadFi x Controle", page_icon="📊", layout="centered")

st.title("📊 1° - Batimento de Fundos — CadFi x Controle")
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
        with st.spinner("Processando arquivos..."):
            # Carrega e prepara
            cadfi_raw = carregar_excel(cadfi_file)
            controle_raw = carregar_excel(controle_file)

            cadfi_filtrado = filtrar_cadfi(cadfi_raw)
            controle_prep = carregar_controle(controle_raw)

            # Comparações
            df_fora = comparar_cnpjs(cadfi_filtrado, controle_prep)                         # CadFi -> não no Controle
            df_comum = comparar_fundos_em_comum(cadfi_filtrado, controle_prep)              # Interseção
            df_controle_fora = comparar_controle_fora_cadfi(cadfi_filtrado, controle_prep)  # Controle -> não no CadFi

            # NOVO 1: remove por SITUACAO ('I' e 'P')
            df_controle_fora = filtrar_controle_por_situacao(df_controle_fora)

            # NOVO 2: remove pelos dois nomes específicos
            df_controle_fora = filtrar_controle_por_nome(df_controle_fora)

            # Relatórios prontos
            rel_fora = relatorio_fora_controle(df_fora)
            rel_comum = relatorio_em_comum(df_comum)
            rel_controle_fora = relatorio_controle_fora_cadfi(df_controle_fora)



        # Contagens
        st.success(f"✅ Em comum: {len(rel_comum)} fundo(s)")
        st.info(f"ℹ️ No Controle e NÃO no CadFi: {len(rel_controle_fora)} fundo(s)")
        st.warning(f"❌ Fora do Controle (presentes no CadFi, ausentes no Controle): {len(rel_fora)} fundo(s)")


        with st.expander("✅ Fundos presentes em AMBOS (CadFi e Controle)"):
            st.dataframe(rel_comum, use_container_width=True, hide_index=True)

        with st.expander("ℹ️ Fundos do Controle que NÃO estão no CadFi"):
            st.dataframe(rel_controle_fora, use_container_width=True, hide_index=True)
            
        with st.expander("❌ Fundos do CadFi que NÃO estão no Controle"):
            st.dataframe(rel_fora, use_container_width=True, hide_index=True)



        # Downloads
        c1, c2, c3 = st.columns(3)
        with c1:
            st.download_button(
                label="⬇️ Baixar — Fundos em AMBOS (CadFi e Controle)",
                data=to_excel_bytes(rel_comum, sheet_name="Em_Comum"),
                file_name="Relatorio_Fundos_Em_Ambos.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )
        with c2:
            st.download_button(
                label="⬇️ Baixar — Somente no CadFi (não no Controle)",
                data=to_excel_bytes(rel_fora, sheet_name="Somente_no_CadFi"),
                file_name="Relatorio_Fundos_Somente_no_CadFi.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )
        with c3:
            st.download_button(
                label="⬇️ Baixar — Somente no Controle (não no CadFi)",
                data=to_excel_bytes(rel_controle_fora, sheet_name="Somente_no_Controle"),
                file_name="Relatorio_Fundos_Somente_no_Controle.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )

    except Exception as e:
        st.exception(e)

st.markdown("---")
st.caption("Dica: se as colunas do Excel vierem com acentos/variações, o app normaliza nomes para evitar erros.")

# ========================== INTERFACE: CDA (Enriquecer "Em Ambos") ==========================
st.markdown("---")
st.subheader("📄 CDA — Enriquecer o relatório **Fundos em Ambos** com Protocolo/Competência")

col_cda1, col_cda2 = st.columns(2)
with col_cda1:
    rel_ambos_file = st.file_uploader("Relatório — Fundos em Ambos (xlsx)", type=["xlsx"], key="rel_ambos_cda")
with col_cda2:
    cda_proto_file = st.file_uploader("Planilha de Protocolo do CDA (xlsx)", type=["xlsx"], key="cda_proto_file")

bt_cda = st.button("Preencher colunas do CDA", type="primary", key="btn_cda_process")

if bt_cda:
    if not rel_ambos_file or not cda_proto_file:
        st.error("⚠️ Envie **os dois arquivos**: (1) Relatório 'Em Ambos' e (2) Protocolo do CDA.")
        st.stop()
    try:
        with st.spinner("Lendo arquivos e integrando CDA..."):
            # Carrega o relatório 'Em Ambos'
            df_ambos = pd.read_excel(rel_ambos_file, dtype=str)
            df_ambos = padronizar_colunas(df_ambos)

            if "CNPJ" not in df_ambos.columns:
                st.error("O relatório 'Em Ambos' precisa ter a coluna 'CNPJ'.")
                st.stop()

            # Garante colunas esperadas (caso venham ausentes)
            if "Protocolo" not in df_ambos.columns:
                df_ambos["Protocolo"] = ""
            if "Mes de Referencia" not in df_ambos.columns:
                df_ambos["Mes de Referencia"] = ""

            # Carrega e parseia o CDA
            df_cda = parse_protocolos_cda_xlsx(cda_proto_file)

            # Enriquecer
            df_final = enriquecer_em_comum_com_cda(df_ambos, df_cda)

            # Métricas
            tot = len(df_final)
            casados = df_final["CDA_Protocolo"].astype(str).str.strip().ne("Não possui").sum()
            st.success(f"✅ Encontramos protocolo do CDA para {casados} de {tot} fundos.")

            with st.expander("🔎 Prévia do relatório enriquecido"):
                st.dataframe(df_final, use_container_width=True, hide_index=True)

            st.download_button(
                label="⬇️ Baixar — 'Em Ambos' enriquecido com CDA",
                data=to_excel_bytes(df_final, sheet_name="Em_Ambos_com_CDA"),
                file_name="Relatorio_Em_Ambos_com_CDA.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )
    except Exception as e:
        st.exception(e)
