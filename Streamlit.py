def parse_protocolo_balancete(arquivo_excel) -> pd.DataFrame:
    df_raw = pd.read_excel(arquivo_excel, sheet_name=0, header=None, dtype=str)
    linhas = []
    for r, row in df_raw.iterrows():
        for c, val in enumerate(row):
            if pd.notna(val) and str(val).strip():
                linhas.append((r, c, str(val).strip()))

    registros = []
    for i, (r, c, txt) in enumerate(linhas):
        txt_up = txt.strip().upper()

        # Se for PARTICIPANTE, tenta capturar o CNPJ e os dados relacionados
        if txt_up.startswith("PARTICIPANTE"):
            cnpj_fmt, cnpj_num = None, None

            # procurar CNPJ logo abaixo
            for rr in range(r, r+6):
                row_vals = df_raw.iloc[rr].astype(str).tolist()
                for cell in row_vals:
                    m = re.search(r"(\d{2}\.\d{3}\.\d{3}/\d{4}-\d{2})", str(cell))
                    if m:
                        cnpj_fmt = m.group(1)
                        cnpj_num = normaliza_cnpj(cnpj_fmt)
                        break
                if cnpj_fmt:
                    break

            # procurar protocolo nas proximidades
            protocolo = None
            for rr in range(max(0, r-10), min(len(df_raw), r+10)):
                row_vals = df_raw.iloc[rr].astype(str).tolist()
                for cell in row_vals:
                    m_proto = re.search(r"N[º°]?\s*PROTOCOLO[:\-\s]*([A-Z0-9\-]{4,30})", str(cell), flags=re.I)
                    if m_proto:
                        protocolo = m_proto.group(1).strip()
                        break
                if protocolo:
                    break

            # procurar competência ANTES do participante
            competencia = None
            for rr in range(max(0, r-15), r+1):
                row_vals = df_raw.iloc[rr].tolist()
                for cc, cell in enumerate(row_vals):
                    if isinstance(cell, str) and cell.strip().upper().startswith("COMPETÊNCIA"):
                        # pega a célula vizinha (mesma linha, próxima coluna)
                        if cc + 1 < len(row_vals):
                            competencia_raw = str(row_vals[cc+1]).strip()
                            # tenta normalizar
                            m = re.search(r"(\d{4})-(\d{2})", competencia_raw)
                            if m:
                                competencia = f"{m.group(2)}/{m.group(1)}"
                            else:
                                competencia = competencia_raw
                        break
                if competencia:
                    break

            if cnpj_num:
                registros.append({
                    "CNPJ": formatar_cnpj(cnpj_num),
                    "Balancete_Protocolo": protocolo or "",
                    "Balancete_Competencia": competencia or ""
                })

    if not registros:
        return pd.DataFrame(columns=["CNPJ", "Balancete_Protocolo", "Balancete_Competencia"])

    df = pd.DataFrame(registros).drop_duplicates(subset="CNPJ", keep="first").reset_index(drop=True)
    return df
