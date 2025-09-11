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
            for j in range(i + 1, min(i + 8, n)):
                m = re.search(r"\((\d{2}\.\d{3}\.\d{3}/\d{4}-\d{2})\)", linhas[j])
                if m:
                    cnpj_fmt = m.group(1)
                    break
            cnpj_num = normaliza_cnpj(cnpj_fmt) if cnpj_fmt else None

            # janela: só linhas ANTES do Participante (até 15 acima)
            ws = max(0, i - 15)
            window = linhas[ws:i]

            # --- Competência: procurar explicitamente "Competência:"
            competencia_raw = None
            for txt in reversed(window):  # de baixo pra cima (linha mais próxima do Participante tem prioridade)
                m = re.search(r"COMPET[ÊE]NCIA\s*:\s*(.+)", txt, flags=re.I)
                if m:
                    competencia_raw = m.group(1).strip()
                    break

            competencia = _normalize_competencia_to_mm_yyyy(competencia_raw)

            # --- Protocolo (mantida a lógica atual)
            ws = max(0, i - 15)
            we = min(n, i + 15)
            window = linhas[ws:we]

            protocolo = None
            # 1) procura "Nº Protocolo: XXXX" inline
            for idx, txt in enumerate(window):
                m_inline = re.search(r"N[\sº°oO]*\s*PROTOCOLO[:\-\s]*([A-Z0-9\-]{4,60})", txt, flags=re.I)
                if m_inline:
                    protocolo = m_inline.group(1).strip()
                    break

            # 2) se não achou, procura linha com "PROTOCOLO" e pega a próxima
            if not protocolo:
                for idx, txt in enumerate(window):
                    if re.search(r"N[\sº°oO]*\s*PROTOCOLO", txt, flags=re.I):
                        if idx + 1 < len(window):
                            cand = window[idx + 1].strip()
                            m_cand = re.search(r"([A-Z0-9\-]{4,60})", cand, flags=re.I)
                            if m_cand:
                                protocolo = m_cand.group(1).strip()
                                break

            # 3) fallback: token alfanumérico próximo
            if not protocolo:
                for idx in range(ws, we):
                    m_any = re.search(r"\b([A-Z0-9]{6,60})\b", linhas[idx], flags=re.I)
                    if m_any:
                        tok = m_any.group(1)
                        if not re.match(r"\d{2}/\d{4}", tok) and not re.match(r"\d{2}/\d{2}/\d{4}", tok):
                            protocolo = tok
                            break

            # grava o registro (se achou CNPJ)
            if cnpj_num:
                registros.append({
                    "CNPJ": formatar_cnpj(cnpj_num),
                    "Balancete_Protocolo": protocolo or "",
                    "Balancete_Competencia": competencia or ""
                })

            # avança pra depois do bloco do participante
            i = (j + 1) if j and (j + 1) > i else i + 1
            continue

        i += 1

    if not registros:
        return pd.DataFrame(columns=["CNPJ", "Balancete_Protocolo", "Balancete_Competencia"])

    df = pd.DataFrame(registros).drop_duplicates(subset="CNPJ", keep="first").reset_index(drop=True)
    return df

