# Substitua o conteúdo do seu Streamlit (5).py por este código completo

import io
import re
import unicodedata
from pathlib import Path
import pandas as pd
import streamlit as st
from typing import Optional, Tuple

# === [NOVO BLOCO] Extração de Protocolo e Competência do Balancete ===
from typing import Optional, Tuple
import re

# Mapa de meses PT-BR -> número
MESES_PT = {
    "JAN": 1, "JANEIRO": 1,
    "FEV": 2, "FEVEREIRO": 2,
    "MAR": 3, "MARCO": 3, "MARÇO": 3,
    "ABR": 4, "ABRIL": 4,
    "MAI": 5, "MAIO": 5,
    "JUN": 6, "JUNHO": 6,
    "JUL": 7, "JULHO": 7,
    "AGO": 8, "AGOSTO": 8,
    "SET": 9, "SETEMBRO": 9, "SETEM": 9, "SETEMB": 9,
    "OUT": 10, "OUTUBRO": 10,
    "NOV": 11, "NOVEMBRO": 11,
    "DEZ": 12, "DEZEMBRO": 12,
}

def _format_competencia_yyyy_mm(ano: int, mes: int) -> str:
    mes = max(1, min(12, int(mes)))
    return f"{int(ano):04d}-{mes:02d}"

def _parse_competencia(texto: str) -> Optional[str]:
    T = normaliza_texto(texto)

    # 1) MM/AAAA ou MM-AAAA
    m = re.search(r"\b(\d{1,2})[/\-](\d{4})\b", T)
    if m:
        mes, ano = int(m.group(1)), int(m.group(2))
        if 1 <= mes <= 12:
            return _format_competencia_yyyy_mm(ano, mes)

    # 2) AAAA-MM ou AAAA/MM
    m = re.search(r"\b(\d{4})[/\-](\d{1,2})\b", T)
    if m:
        ano, mes = int(m.group(1)), int(m.group(2))
        if 1 <= mes <= 12:
            return _format_competencia_yyyy_mm(ano, mes)

    # 3) Nome do mês (abreviado ou completo) + AAAA
    m = re.search(r"\b([A-ZÇÃÉ]+)[\s/.\-]*(\d{4})\b", T)
    if m:
        mes_txt, ano = m.group(1), int(m.group(2))
        mes = MESES_PT.get(mes_txt)
        if mes:
            return _format_competencia_yyyy_mm(ano, mes)

    return None

def _eh_cnpj_sequencia(numeros: str) -> bool:
    d = re.sub(r"\D", "", str(numeros or ""))
    return len(d) == 14

def _parse_protocolo(texto: str) -> Optional[str]:
    T = normaliza_texto(texto)

    m = re.search(r"(?:PROTOCOLO|NUMERO\s*DE\s*PROTOCOLO|GFI)\D*(\d{6,})", T, flags=re.I)
    if m:
        valor = m.group(1)
        if not _eh_cnpj_sequencia(valor):
            return valor

    candidatos = re.findall(r"\b(\d{6,})\b", T)
    candidatos = [c for c in candidatos if not _eh_cnpj_sequencia(c)]
    if candidatos:
        return max(candidatos, key=len)

    return None

def _read_text_from_xlsx(uploaded_file) -> str:
    try:
        df_head = pd.read_excel(uploaded_file, header=None, nrows=40, dtype=str, engine="openpyxl")
        texto = " ".join(df_head.astype(str).fillna("").values.ravel())
        return texto
    except Exception:
        return ""
    finally:
        try:
            uploaded_file.seek(0)
        except Exception:
            pass

def _read_text_from_pdf(uploaded_file) -> str:
    try:
        import fitz  # PyMuPDF
    except Exception:
        return ""

    try:
        data = uploaded_file.read()
        texto = ""
        with fitz.open(stream=data, filetype="pdf") as doc:
            for page in doc:
                texto += " " + page.get_text("text")
        return texto
    except Exception:
        return ""
    finally:
        try:
            uploaded_file.seek(0)
        except Exception:
            pass

def extrair_protocolo_e_competencia_do_balancete(uploaded_file) -> Tuple[Optional[str], Optional[str]]:
    if not uploaded_file:
        return (None, None)

    nome = str(getattr(uploaded_file, "name", "")).lower()
    texto = ""

    if nome.endswith(".xlsx"):
        texto = _read_text_from_xlsx(uploaded_file)
    elif nome.endswith(".pdf"):
        texto = _read_text_from_pdf(uploaded_file)

    texto_total = f"{texto}  {getattr(uploaded_file, 'name', '')}"

    protocolo = _parse_protocolo(texto_total)
    competencia = _parse_competencia(texto_total)

    return (protocolo, competencia)
# === [FIM DO BLOCO NOVO] ===

def so_digitos(s):
    return re.sub(r'\D', '', str(s or ''))

def normaliza_cnpj(cnpj):
    d = so_digitos(cnpj)
    if len(d) == 14:
        return d
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
    norm_map = {_norm_header_key(c): c for c in df.columns}
    prioridade = [
        "situacao", "situação", "situacao_do_fundo", "situcao", "status", "status_do_fundo"
    ]
    for key in prioridade:
        if key in norm_map:
            return norm_map[key]
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

VALORES_ATIVOS = {
    normaliza_texto("Em Funcionamento Normal"),
    normaliza_texto("Em Funcionamento"),
    normaliza_texto("Ativo"),
    normaliza_texto("Ativa"),
    normaliza_texto("Em Atividade"),
    normaliza_texto("A"),
}

def filtrar_status_ativos(df: pd.DataFrame) -> pd.DataFrame:
    if df is None or df.empty:
        return df
    col = _encontrar_coluna_status(df)
    if not col:
        return df
    out = df.copy()
    out["_STATUS_NORM_"] = out[col].map(normaliza_texto)
    out = out[out["_STATUS_NORM_"].isin(VALORES_ATIVOS)].drop(columns=["_STATUS_NORM_"])
    return out

def carregar_excel(arquivo):
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
        (df["Tipo_Fundo"].isin(["FI", "FAPI", "FIIM"])) &
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
    return controle_df[~controle_df["CNPJ"].isin(set(cadfi_df["CNPJ"]))].copy()

def _encontrar_coluna_nome(df: pd.DataFrame) -> str:
    norm_map = {_norm_header_key(c): c for c in df.columns}
    prioridade = [
        "denominacao_social", "denominacao_do_fundo", "denominacao",
        "nome_do_fundo", "nome_fundo", "nome",
        "razao_social", "razao", "descricao"
    ]
    for key in prioridade:
        if key in norm_map:
            return norm_map[key]
    candidatos = []
    for k, original in norm_map.items():
        score = 0
        if "denomin" in k: score += 3
        if "nome" in k:    score += 2
        if "fundo" in k:   score += 1
        if "cnpj" in k:    score = -1
        if score > 0:
            candidatos.append((score, original))
    if candidatos:
        candidatos.sort(reverse=True, key=lambda x: x[0])
        return candidatos[0][1]
    for c in df.columns:
        if c != "CNPJ" and df[c].dtype == object:
            return c
    return None

def relatorio_controle_fora_cadfi(df_controle: pd.DataFrame) -> pd.DataFrame:
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
        out["Nome do fundo (Controle)"] = ""
    return out[["CNPJ", "Nome do fundo (Controle)"]]

EXCLUIR_NOMES_CONTROLE = [
    "BB CIN",
    "BB BNC AÇÕES NOSSA CAIXA NOSSO CLUBE DE INVESTIMENTO",
]

def filtrar_controle_por_nome(df: pd.DataFrame,
                              nomes_excluir=EXCLUIR_NOMES_CONTROLE) -> pd.DataFrame:
    if df is None or df.empty:
        return df
    col_nome = _encontrar_coluna_nome(df)
    if not col_nome or col_nome not in df.columns:
        return df
    nomes_norm = [normaliza_texto(n) for n in nomes_excluir]
    out = df.copy()
    out["_NOME_NORM_"] = out[col_nome].map(normaliza_texto)
    mask_excluir = out["_NOME_NORM_"].apply(lambda s: any(p in s for p in nomes_norm))
    out = out[~mask_excluir].drop(columns=["_NOME_NORM_"])
    return out

EXCLUIR_SITUACAO_CONTROLE = ("I", "P")

def filtrar_controle_por_situacao(df: pd.DataFrame,
                                  excluir_codigos=EXCLUIR_SITUACAO_CONTROLE) -> pd.DataFrame:
    if df is None or df.empty:   # ✅ corrigido
        return df

    col_status = _encontrar_coluna_status(df)
    if not col_status or col_status not in df.columns:
        return df

    excluir_norm = {normaliza_texto(x)[:1] for x in excluir_codigos}
    out = df.copy()
    out["_SIT_"] = out[col_status].map(
        lambda x: normaliza_texto(x)[:1] if pd.notna(x) else ""
    )
    mask_excluir = out["_SIT_"].isin(excluir_norm)
    out = out[~mask_excluir].drop(columns=["_SIT_"])
    return out


def carregar_controle(df_controle):
    if "CNPJ" not in df_controle.columns:
        raise ValueError("Coluna 'CNPJ' ausente no Controle Espelho.")
    return remover_duplicatas_por_cnpj(df_controle, "CNPJ")

def comparar_cnpjs(cadfi_df, controle_df):
    return cadfi_df[~cadfi_df["CNPJ"].isin(set(controle_df["CNPJ"]))].copy()

def comparar_fundos_em_comum(cadfi_df, controle_df):
    return cadfi_df[cadfi_df["CNPJ"].isin(set(controle_df["CNPJ"]))].copy()

def relatorio_fora_controle(df):
    if df is None or df.empty:
        return pd.DataFrame(columns=["CNPJ", "Nome do fundo"])
    df = df.copy()
    rel = df[["CNPJ", "Denominacao_Social"]].rename(columns={
        "Denominacao_Social": "Nome do fundo",
    })
    return rel

def relatorio_em_comum(df):
    if df is None or df.empty:
        return pd.DataFrame(columns=[
            "CNPJ", "Nome do fundo"
        ])
    df = df.copy()
    rel = df[[
        "CNPJ", "Denominacao_Social"
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

def _normaliza_competencia_mm_aaaa(s: str) -> Optional[str]:
    if not s:
        return None
    s = s.strip()
    m_iso = re.search(r'(20\d{2})[/\-](\d{2})', s)
    if m_iso:
        ano, mes = int(m_iso.group(1)), int(m_iso.group(2))
        if 1 <= mes <= 12:
            return _format_competencia_yyyy_mm(ano, mes)
    m_br = re.search(r'(\d{2})[/\-](20\d{2})', s)
    if m_br:
        mes, ano = int(m_br.group(1)), int(m_br.group(2))
        if 1 <= mes <= 12:
            return _format_competencia_yyyy_mm(ano, mes)
    return None

def remover_segundos_colunas(df: pd.DataFrame, colunas, formato: str = "%Y-%m-%d %H:%M") -> pd.DataFrame:
    df = df.copy()
    for col in colunas:
        if col in df.columns:
            s = pd.to_datetime(df[col], errors="coerce")
            df.loc[s.notna(), col] = s[s.notna()].dt.strftime(formato)
            df.loc[s.isna(), col] = (
                df.loc[s.isna(), col]
                .astype(str)
                .str.replace(r":\d{2}(?=\b)", "", regex=True)
            )
    return df

def parse_protocolos_cda_xlsx(arquivo_xlsx) -> pd.DataFrame:
    df_raw = pd.read_excel(arquivo_xlsx, sheet_name=0, header=None, dtype=str)
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
        if low.startswith('nº protocolo') or low.startswith('n° protocolo') or low.startswith('no protocolo') \
           or low.startswith('nº do protocolo') or low.startswith('n° do protocolo'):
            protocolo = None
            j = i + 1
            while j < n:
                _, t2 = lines[j]
                low2 = t2.lower()
                if re.match(
                    r'^(protocolo de confirm|status:|informe:|opera|documento:|compet|usuário|usuario|nº do recebimento|nome do arquivo|participante:|tipo do participante|data ação:|data acao:)',
                    low2
                ):
                    j += 1
                    continue
                protocolo = t2.strip()
                if protocolo.endswith(".0"):
                    protocolo = protocolo[:-2]
                break
                j += 1
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
                        if low2.startswith('tipo do participante') or low2.startswith('data ação') or low2.startswith('data acao') or low2.startswith('nº protocolo') or low2.startswith('n° protocolo') or low2.startswith('nº do protocolo') or low2.startswith('n° do protocolo'):
                            break
                        if first_name is None and tline:
                            first_name = tline.strip()
                        m = re.search(r'(\d{2}\.\d{3}\.\d{3}/\d{4}-\d{2})', tline)
                        if m:
                            cnpj_masked = m.group(1)
                            break
                        kk += 1
                    participante = first_name
                    break
                k -= 1
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
            if cnpj_masked and protocolo:
                cnpj_num = normaliza_cnpj(cnpj_masked)
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
    df = df.sort_values(["CNPJ_Num", "Data_Acao"], ascending=[True, False]).drop_duplicates("CNPJ_Num", keep="first")
    return df

def enriquecer_em_comum_com_cda(rel_em_comum_df: pd.DataFrame, df_cda: pd.DataFrame) -> pd.DataFrame:
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
    cols = list(enx.columns)
    insert_pos = cols.index("Mes de Referencia") + 1 if "Mes de Referencia" in cols else len(cols)
    for newcol in ["CDA_Protocolo", "CDA_Competencia"]:
        if newcol in cols:
            cols.remove(newcol)
    cols = cols[:insert_pos] + ["CDA_Protocolo", "CDA_Competencia"] + cols[insert_pos:]
    enx = enx[cols]
    if "CNPJ_Num" in enx.columns:
        enx = enx.drop(columns=["CNPJ_Num"])
    return enx

# ======================= /CDA =====================================================

# ======================== Balancete ===============================================

import itertools

def _linhas_excel_como_texto(arquivo_excel) -> list[str]:
    xls = pd.ExcelFile(arquivo_excel, engine="openpyxl")
    linhas = []
    for sheet in xls.sheet_names:
        df = pd.read_excel(xls, sheet_name=sheet, dtype=str, header=None, engine="openpyxl")
        for _, row in df.iterrows():
            for val in row.tolist():
                s = str(val).strip() if pd.notna(val) else ""
                if s:
                    linhas.append(s)
    return [unicodedata.normalize("NFKD", s).strip() for s in linhas if s.strip()]

def _extrair_mm_yyyy_de_nome_arquivo(linhas: list[str]) -> Optional[str]:
    mm_yyyy = None
    for i, text in enumerate(linhas):
        if text.upper().startswith("NOME DO ARQUIVO"):
            for j in range(i+1, min(i+5, len(linhas))):
                cand = linhas[j]
                m = re.search(r"(\d{6})(?!\d)", cand)
                if m:
                    mm = m.group(1)[:2]
                    yyyy = m.group(1)[2:]
                    mm_yyyy = f"{mm}/{yyyy}"
                    return mm_yyyy
    return mm_yyyy

def parse_protocolo_balancete(arquivo_excel) -> pd.DataFrame:
    linhas = _linhas_excel_como_texto(arquivo_excel)
    registros = []
    n = len(linhas)
    i = 0

    while i < n:
        s = linhas[i].strip().upper()
        if s.startswith("PARTICIPANTE"):
            # procura CNPJ logo abaixo do "Participante:"
            j = i
            cnpj_fmt = None
            for j in range(i+1, min(i+8, n)):
                m = re.search(r"\((\d{2}\.\d{3}\.\d{3}/\d{4}-\d{2})\)", linhas[j])
                if m:
                    cnpj_fmt = m.group(1)
                    break
            cnpj_num = normaliza_cnpj(cnpj_fmt) if cnpj_fmt else None

            # janela de busca ao redor do bloco (antes e depois) para protocolo/competência
            ws = max(0, i-15)
            we = min(n, i+15)
            window = linhas[ws:we]

            protocolo = None
            # 1) procura "Nº Protocolo: XXXX" inline (alfa-numérico)
            for idx, txt in enumerate(window):
                m_inline = re.search(r"N[\sº°oO]*\s*PROTOCOLO[:\-\s]*([A-Z0-9\-]{4,30})", txt, flags=re.I)
                if m_inline:
                    protocolo = m_inline.group(1).strip()
                    break

            # 2) se não achou, procura linha contendo "PROTOCOLO" e pega próxima linha como token
            if not protocolo:
                for idx, txt in enumerate(window):
                    if re.search(r"N[\sº°oO]*\s*PROTOCOLO", txt, flags=re.I):
                        if idx + 1 < len(window):
                            cand = window[idx+1].strip()
                            m_cand = re.search(r"([A-Z0-9\-]{4,30})", cand, flags=re.I)
                            if m_cand:
                                protocolo = m_cand.group(1).strip()
                                break

            # 3) fallback: achar qualquer token alfanumérico razoável próximo (evita datas/CNPJs)
            if not protocolo:
                for idx in range(ws, we):
                    m_any = re.search(r"\b([A-Z0-9]{6,20})\b", linhas[idx], flags=re.I)
                    if m_any:
                        tok = m_any.group(1)
                        if not re.match(r"\d{2}/\d{4}", tok) and not re.match(r"\d{2}/\d{2}/\d{4}", tok):
                            protocolo = tok
                            break

            # --- Competência: procura dd/mm/yyyy (preferível) ou mm/yyyy no mesmo bloco
            competencia = None
            best = None  # (dist, matchobj)
            for idx_off, txt in enumerate(window):
                m1 = re.search(r"(\d{2})/(\d{2})/(\d{4})", txt)
                if m1:
                    abspos = ws + idx_off
                    dist = abs(abspos - i)
                    if best is None or dist < best[0]:
                        best = (dist, m1)
            if best:
                m = best[1]
                competencia = f"{m.group(2)}/{m.group(3)}"  # MM/AAAA
            else:
                # procura mm/yyyy
                for txt in window:
                    m2 = re.search(r"\b(\d{2})/(\d{4})\b", txt)
                    if m2:
                        competencia = f"{m2.group(1)}/{m2.group(2)}"
                        break

            if cnpj_num:
                registros.append({
                    "CNPJ": formatar_cnpj(cnpj_num),
                    "Balancete_Protocolo": protocolo or "",
                    "Balancete_Competencia": competencia or ""
                })

            # avança o índice para sair do bloco (se achou j, vai pra depois dele)
            i = (j + 1) if j and (j + 1) > i else i + 1
            continue

        i += 1

    if not registros:
        return pd.DataFrame(columns=["CNPJ", "Balancete_Protocolo", "Balancete_Competencia"])

    df = pd.DataFrame(registros).drop_duplicates(subset="CNPJ", keep="first").reset_index(drop=True)
    return df


def parse_protocolo_balancete_from_pdf(uploaded_pdf) -> pd.DataFrame:
    text = _read_text_from_pdf(uploaded_pdf)
    if not text:
        return pd.DataFrame(columns=["CNPJ", "Balancete_Protocolo", "Balancete_Competencia"])
    # Heurística: encontre CNPJs e, para cada um, busque protocolo nas proximidades
    pattern_cnpj = re.compile(r"(\d{2}\.\d{3}\.\d{3}/\d{4}-\d{2})")
    pattern_proto = re.compile(r"(?:N[º°]\s*PROTOCOLO|PROTOCOLO)\D*(\d{6,})", flags=re.I)
    pattern_comp_mm_yyyy = re.compile(r"(\b\d{2}/\d{4}\b)")
    registros = []
    for m in pattern_cnpj.finditer(text):
        cnpj_masked = m.group(1)
        start = m.start()
        # busca protocolo nos 500 caracteres seguintes
        window = text[start:start+500]
        proto_m = pattern_proto.search(window)
        protocolo = proto_m.group(1) if proto_m else ""
        # busca competência no trecho anterior (200 chars)
        prev_window = text[max(0, start-400):start+200]
        comp_m = pattern_comp_mm_yyyy.search(prev_window)
        competencia = comp_m.group(1) if comp_m else ""
        cnpj_num = normaliza_cnpj(cnpj_masked)
        if cnpj_num:
            registros.append({
                "CNPJ": formatar_cnpj(cnpj_num),
                "Balancete_Protocolo": protocolo,
                "Balancete_Competencia": competencia
            })
    if not registros:
        return pd.DataFrame(columns=["CNPJ", "Balancete_Protocolo", "Balancete_Competencia"])
    df = pd.DataFrame(registros)
    df = df.drop_duplicates(subset="CNPJ", keep="first").reset_index(drop=True)
    return df

# ========================== INTERFACE STREAMLIT ==========================
st.set_page_config(page_title="Batimento de Fundos - CadFi x Controle", page_icon="📊", layout="centered")

st.title("Batimento de Fundos — Contabilidade")
st.subheader("📊 1° - Batimento de Fundos — CadFi x Controle")
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
            cadfi_raw = carregar_excel(cadfi_file)
            controle_raw = carregar_excel(controle_file)

            cadfi_filtrado = filtrar_cadfi(cadfi_raw)
            controle_prep = carregar_controle(controle_raw)

            df_fora = comparar_cnpjs(cadfi_filtrado, controle_prep)
            df_comum = comparar_fundos_em_comum(cadfi_filtrado, controle_prep)
            df_controle_fora = comparar_controle_fora_cadfi(cadfi_filtrado, controle_prep)

            df_controle_fora = filtrar_controle_por_situacao(df_controle_fora)
            df_controle_fora = filtrar_controle_por_nome(df_controle_fora)

            rel_fora = relatorio_fora_controle(df_fora)
            rel_comum = relatorio_em_comum(df_comum)
            rel_comum = remover_segundos_colunas(rel_comum, ["CDA_Protocolo", "CDA_Competencia"])
            rel_controle_fora = relatorio_controle_fora_cadfi(df_controle_fora)

        st.success(f"✅ Em comum: {len(rel_comum)} fundo(s)")
        st.info(f"ℹ️ No Controle e NÃO no CadFi: {len(rel_controle_fora)} fundo(s)")
        st.warning(f"❌ Fora do Controle (presentes no CadFi, ausentes no Controle): {len(rel_fora)} fundo(s)")

        with st.expander("✅ Fundos presentes em AMBOS (CadFi e Controle)"):
            st.dataframe(rel_comum, use_container_width=True, hide_index=True)

        with st.expander("ℹ️ Fundos do Controle que NÃO estão no CadFi"):
            st.dataframe(rel_controle_fora, use_container_width=True, hide_index=True)

        with st.expander("❌ Fundos do CadFi que NÃO estão no Controle"):
            st.dataframe(rel_fora, use_container_width=True, hide_index=True)

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
st.subheader("📄 2° CDA — Enriquecer o relatório **Fundos em Ambos** com Protocolo/Competência")

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
            df_ambos = pd.read_excel(rel_ambos_file, dtype=str)
            df_ambos = padronizar_colunas(df_ambos)

            if "CNPJ" not in df_ambos.columns:
                st.error("O relatório 'Em Ambos' precisa ter a coluna 'CNPJ'.")
                st.stop()

            df_cda = parse_protocolos_cda_xlsx(cda_proto_file)
            df_final = enriquecer_em_comum_com_cda(df_ambos, df_cda)

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

#====================================== Interface de Balancete -============================
st.markdown("## 🔄 3° - Enriquecer batimento com Balancete")
colb1, colb2 = st.columns(2)
with colb1:
    relatorio_ambos_file = st.file_uploader("Arquivo Relatório de Ambos (.xlsx)", type=["xlsx"], key="relatorio_ambos")
with colb2:
    balancete_file = st.file_uploader(
        "Arquivo de Balancete (XLSX ou PDF)",
        type=["xlsx", "pdf"],
        accept_multiple_files=False
    )

enriquecer = st.button("Preencher colunas Balancete", type="primary", key="btn_balancete_enriquecer")

if enriquecer:
    if not relatorio_ambos_file or not balancete_file:
        st.error("⚠️ Envie os dois arquivos antes de enriquecer.")
        st.stop()

    try:
        with st.spinner("Enriquecendo com Balancete..."):
            # Carrega relatório 'Em Ambos'
            df_rel_comum = pd.read_excel(relatorio_ambos_file, dtype=str)
            df_rel_comum = padronizar_colunas(df_rel_comum)

            if "CNPJ" not in df_rel_comum.columns:
                st.error("O relatório 'Em Ambos' precisa ter a coluna 'CNPJ'.")
                st.stop()

            # Parse do arquivo de balancete (xlsx mais confiável; pdf heurístico)
            fname = str(getattr(balancete_file, "name", "")).lower()
            if fname.endswith(".xlsx"):
                df_balancete_proto = parse_protocolo_balancete(balancete_file)
            else:
                df_balancete_proto = parse_protocolo_balancete_from_pdf(balancete_file)

            # Garantir padronização/formatacao das colunas e CNPJ
            df_balancete_proto = padronizar_colunas(df_balancete_proto)
            if "CNPJ" in df_balancete_proto.columns:
                df_balancete_proto["CNPJ"] = df_balancete_proto["CNPJ"].apply(lambda x: formatar_cnpj(normaliza_cnpj(x)) if pd.notna(x) else None)
            # preencher nomes das colunas esperadas se vierem com nomes diferentes
            # já trabalhamos com "Balancete_Protocolo" e "Balancete_Competencia"

            # Merge (normalizando formato CNPJ em ambos)
            df_rel_comum["CNPJ"] = df_rel_comum["CNPJ"].apply(lambda x: formatar_cnpj(normaliza_cnpj(x)) if pd.notna(x) else None)
            if "CNPJ" not in df_balancete_proto.columns:
                st.warning("Não foi possível extrair CNPJ do arquivo de Balancete — verifique o layout. Resultado pode ficar vazio.")
            merged = df_rel_comum.merge(
                df_balancete_proto[["CNPJ", "Balancete_Protocolo", "Balancete_Competencia"]],
                on="CNPJ",
                how="left"
            )

            merged["Balancete_Protocolo"] = merged["Balancete_Protocolo"].fillna("Não possui")
            merged["Balancete_Competencia"] = merged["Balancete_Competencia"].fillna("Não possui")

            # Posicionar colunas após 'Mes de Referencia' se existir
            cols = list(merged.columns)
            insert_pos = cols.index("Mes de Referencia") + 1 if "Mes de Referencia" in cols else len(cols)
            for newcol in ["Balancete_Protocolo", "Balancete_Competencia"]:
                if newcol in cols:
                    cols.remove(newcol)
            cols = cols[:insert_pos] + ["Balancete_Protocolo", "Balancete_Competencia"] + cols[insert_pos:]
            merged = merged[cols]

            # Mostrar resultado e download
            encontrados = merged["Balancete_Protocolo"].astype(str).str.strip().ne("Não possui").sum()
            st.success(f"✅ Enriquecido com {encontrados} protocolos encontrados.")
            st.dataframe(merged, use_container_width=True, hide_index=True)

            st.download_button(
                label="⬇️ Baixar — Relatório Enriquecido com Balancete",
                data=to_excel_bytes(merged, sheet_name="Enriquecido_Balancete"),
                file_name="Relatorio_Enriquecido_Balancete.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )
    except Exception as e:
        st.exception(e)
