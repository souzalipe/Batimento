def parse_protocolo_balancete(arquivo_excel) -> pd.DataFrame:
    """
    Extração robusta do Nº do Protocolo do Balancete por bloco 'Participante'.
    Regras:
    - bloco = da linha do 'Participante' até próxima ocorrência de 'Participante' ou +60 linhas
    - procura rótulos contendo 'PROTOCOLO' (não ancorado)
      -> checa mesma linha: próximas 1..3 células à direita
      -> checa linha abaixo: mesma coluna
    - fallback: primeiro 'SCW\d{6,}' dentro do bloco (pra frente)
    - detecção simples de ETF: se houver 'ETF' no bloco, marca "Não possui"
    - retorna DataFrame com colunas: CNPJ, Balancete_Protocolo, Balancete_Competencia
    """
    df_raw = pd.read_excel(arquivo_excel, sheet_name=0, header=None, dtype=str)
    registros = []
    upper = df_raw.shape[0]
    ncols = df_raw.shape[1] if df_raw.shape[1] > 0 else 0

    for r in range(upper):
        row_vals = df_raw.iloc[r].astype(str).tolist()
        for c, cell in enumerate(row_vals):
            if not isinstance(cell, str):
                continue
            if cell.strip().upper().startswith("PARTICIPANTE"):
                # define limite do bloco: próximo 'PARTICIPANTE' ou +60 linhas
                limite = min(r + 60, upper)
                for rr2 in range(r + 1, min(upper, r + 61)):
                    line_vals = df_raw.iloc[rr2].astype(str).tolist()
                    if any('PARTICIPANTE' in str(x).strip().upper() for x in line_vals):
                        limite = rr2
                        break

                # captura CNPJ dentro do bloco (primeiro que aparecer)
                cnpj_fmt = None
                for rr in range(r, limite):
                    for cell2 in df_raw.iloc[rr].astype(str).tolist():
                        m = re.search(r"(\d{2}\.\d{3}\.\d{3}/\d{4}-\d{2})", str(cell2))
                        if m:
                            cnpj_fmt = m.group(1)
                            break
                    if cnpj_fmt:
                        break
                cnpj_num = normaliza_cnpj(cnpj_fmt) if cnpj_fmt else None

                # detecta ETF simples (se houver a palavra ETF no bloco)
                is_etf = False
                for rr in range(r, limite):
                    if any(re.search(r'\bETF\b', str(x), flags=re.I) for x in df_raw.iloc[rr].astype(str).tolist()):
                        is_etf = True
                        break

                protocolo = None

                if is_etf:
                    protocolo = "Não possui"
                else:
                    # 1) heurística rótulo -> valor (varre só para frente dentro do bloco)
                    encontrado = False
                    for rr in range(r, limite):
                        row_vals2 = df_raw.iloc[rr].astype(str).tolist()
                        for cc, cell2 in enumerate(row_vals2):
                            lab = str(cell2).strip().upper()
                            if 'PROTOCOLO' in lab:
                                # mesma linha: checa até +3 células à direita
                                for k in range(cc + 1, min(cc + 4, len(row_vals2))):
                                    try:
                                        cand_cell = row_vals2[k]
                                    except Exception:
                                        cand_cell = ""
                                    m = re.search(r'(SCW\d{6,})', str(cand_cell), flags=re.I)
                                    if m:
                                        protocolo = m.group(1).upper()
                                        encontrado = True
                                        break
                                if encontrado:
                                    break
                                # linha de baixo: mesma coluna
                                if rr + 1 < upper and cc < ncols:
                                    try:
                                        below = df_raw.iat[rr + 1, cc]
                                    except Exception:
                                        below = ""
                                    m = re.search(r'(SCW\d{6,})', str(below), flags=re.I)
                                    if m:
                                        protocolo = m.group(1).upper()
                                        encontrado = True
                                        break
                        if encontrado:
                            break

                    # 2) fallback: primeiro SCW dentro do bloco (pra frente)
                    if not protocolo:
                        for rr in range(r, limite):
                            for cell2 in df_raw.iloc[rr].astype(str).tolist():
                                m = re.search(r'(SCW\d{6,})', str(cell2), flags=re.I)
                                if m:
                                    protocolo = m.group(1).upper()
                                    break
                            if protocolo:
                                break

                # captura competência (procura para trás até 15 linhas)
                competencia = None
                for rr in range(max(0, r - 15), r + 1):
                    row_vals3 = df_raw.iloc[rr].astype(str).tolist()
                    for cc, cell3 in enumerate(row_vals3):
                        if isinstance(cell3, str) and 'COMPET' in cell3.strip().upper():
                            # tenta célula à direita imediata
                            val = None
                            if cc + 1 < len(row_vals3) and str(row_vals3[cc + 1]).strip():
                                val = row_vals3[cc + 1]
                            else:
                                # tenta na linha de baixo mesma coluna (até 3 linhas)
                                for kk in range(rr + 1, min(rr + 4, upper)):
                                    if cc < df_raw.shape[1]:
                                        v = df_raw.iat[kk, cc]
                                        if pd.notna(v) and str(v).strip():
                                            val = v
                                            break
                            if val:
                                try:
                                    # tenta normalizar para YYYY-MM (compatível com outras funções do script)
                                    comp_norm = _normaliza_competencia_mm_aaaa(str(val).strip())
                                    competencia = comp_norm or str(val).strip()
                                except Exception:
                                    competencia = str(val).strip()
                            break
                    if competencia:
                        break

                registros.append({
                    "CNPJ": formatar_cnpj(cnpj_num) if cnpj_num else None,
                    "Balancete_Protocolo": protocolo or "",
                    "Balancete_Competencia": competencia or ""
                })

    if not registros:
        return pd.DataFrame(columns=["CNPJ", "Balancete_Protocolo", "Balancete_Competencia"])

    df = pd.DataFrame(registros).drop_duplicates(subset="CNPJ", keep="first").reset_index(drop=True)
    return df
