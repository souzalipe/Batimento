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
