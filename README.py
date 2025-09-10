def parse_protocolo_balancete(arquivo_excel) -> pd.DataFrame:
    linhas = _linhas_excel_como_texto(arquivo_excel)
    registros = []
    atual_comp = None

    i = 0
    while i < len(linhas):
        s = linhas[i].strip().upper()

        if s.startswith("PARTICIPANTE"):
            # --- pega CNPJ
            cnpj_fmt = None
            for j in range(i+1, min(i+6, len(linhas))):
                m = re.search(r"\((\d{2}\.\d{3}\.\d{3}/\d{4}-\d{2})\)", linhas[j])
                if m:
                    cnpj_fmt = m.group(1)
                    break

            cnpj_norm = normaliza_cnpj(cnpj_fmt) if cnpj_fmt else None

            # --- pega protocolo (variações Nº, N°, No, etc.)
            protocolo = None
            for k in range(j, min(j+6, len(linhas))):
                if re.search(r"N[\sº°oO]*\s*PROTOCOLO", linhas[k], flags=re.I):
                    if k + 1 < len(linhas):
                        protocolo = linhas[k+1].strip()
                        break

            # --- pega competência (formato dd/mm/yyyy -> mm/yyyy)
            competencia = None
            for k in range(i, min(i+15, len(linhas))):
                if "COMPETÊNCIA" in linhas[k].upper():
                    if k + 1 < len(linhas):
                        raw = linhas[k+1].strip()
                        m_comp = re.match(r"(\d{2})/(\d{2})/(\d{4})", raw)
                        if m_comp:
                            competencia = f"{m_comp.group(2)}/{m_comp.group(3)}"  # MM/AAAA
                        break

            if cnpj_norm:
                registros.append({
                    "CNPJ": formatar_cnpj(cnpj_norm),
                    "Balancete_Protocolo": protocolo or "Não possui",
                    "Balancete_Competencia": competencia or "Não possui"
                })

        i += 1

    if not registros:
        return pd.DataFrame(columns=["CNPJ", "Balancete_Protocolo", "Balancete_Competencia"])

    df = pd.DataFrame(registros).drop_duplicates(subset="CNPJ", keep="first").reset_index(drop=True)
    return df
