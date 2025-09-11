# --- Helper robusto para normalizar competência para MM/YYYY
def _normalize_competencia_to_mm_yyyy(raw: Optional[str]) -> Optional[str]:
    if not raw:
        return None
    s = normaliza_texto(raw).replace(".", "/").replace("-", "/").replace("\\", "/")
    s = s.replace("  ", " ").strip()

    # 1) dd/mm/yyyy -> MM/YYYY
    m = re.search(r"(\d{2})/(\d{2})/(\d{4})", s)
    if m:
        dd, mm, yyyy = m.group(1), m.group(2), m.group(3)
        try:
            mm_i = int(mm)
            if 1 <= mm_i <= 12:
                return f"{mm_i:02d}/{int(yyyy)}"
        except Exception:
            pass

    # 2) mm/yyyy -> MM/YYYY
    m = re.search(r"\b(\d{1,2})/(\d{4})\b", s)
    if m:
        mm, yyyy = int(m.group(1)), int(m.group(2))
        if 1 <= mm <= 12:
            return f"{mm:02d}/{yyyy}"

    # 3) abreviação/nome do mês + ano (ex: jun/25, junho/25, jun/2025, JUN/25)
    m = re.search(r"\b([A-ZÇÃÉÀ-ÿ]{3,10})[^\dA-Z]*(\d{2}|\d{4})\b", s, flags=re.I)
    if m:
        mes_txt = normaliza_texto(m.group(1))
        # tenta mapear a palavra inteira, depois os 3 primeiros chars
        mes_num = MESES_PT.get(mes_txt) or MESES_PT.get(mes_txt[:3]) if mes_txt else None
        if mes_num:
            ano_raw = m.group(2)
            ano = int(ano_raw) + 2000 if len(ano_raw) == 2 else int(ano_raw)
            if 1 <= mes_num <= 12:
                return f"{mes_num:02d}/{ano}"

    # 4) AAAA-MM ou AAAA/MM -> MM/YYYY
    m = re.search(r"\b(20\d{2})[\/\-](\d{1,2})\b", s)
    if m:
        ano, mm = int(m.group(1)), int(m.group(2))
        if 1 <= mm <= 12:
            return f"{mm:02d}/{ano}"

    return None


# --- Substitua sua parse_protocolo_balancete por esta (XLSX)
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
                m_inline = re.search(r"N[\sº°oO]*\s*PROTOCOLO[:\-\s]*([A-Z0-9\-]{4,60})", txt, flags=re.I)
                if m_inline:
                    protocolo = m_inline.group(1).strip()
                    break

            # 2) se não achou, procura linha contendo "PROTOCOLO" e pega próxima linha como token
            if not protocolo:
                for idx, txt in enumerate(window):
                    if re.search(r"N[\sº°oO]*\s*PROTOCOLO", txt, flags=re.I):
                        if idx + 1 < len(window):
                            cand = window[idx+1].strip()
                            m_cand = re.search(r"([A-Z0-9\-]{4,60})", cand, flags=re.I)
                            if m_cand:
                                protocolo = m_cand.group(1).strip()
                                break

            # 3) fallback: achar qualquer token alfanumérico razoável próximo (evita datas/CNPJs)
            if not protocolo:
                for idx in range(ws, we):
                    m_any = re.search(r"\b([A-Z0-9]{6,60})\b", linhas[idx], flags=re.I)
                    if m_any:
                        tok = m_any.group(1)
                        if not re.match(r"\d{2}/\d{4}", tok) and not re.match(r"\d{2}/\d{2}/\d{4}", tok):
                            protocolo = tok
                            break

            # --- Competência: primeiro procura EXPLÍCITO "Competência" no bloco
            competencia_raw = None
            for idx, txt in enumerate(window):
                if "COMPETENCIA" in txt.upper() or "COMPETÊNCIA" in txt.upper():
                    # pega a próxima linha não-vazia se houver
                    if idx + 1 < len(window):
                        cand = window[idx+1].strip()
                        if cand:
                            competencia_raw = cand
                            break
                    # se não tiver próxima linha, tenta extrair da mesma linha
                    m_inline = re.search(r"COMPETENCI[AÉE]\s*[:\-]?\s*(.+)$", txt, flags=re.I)
                    if m_inline:
                        competencia_raw = m_inline.group(1).strip()
                        break

            # se não achou explícito, pega heurística antiga (mais próxima de datas dd/mm/yyyy)
            if not competencia_raw:
                # mantém a heurística anterior mas transformaremos com o normalizador
                best = None
                for idx_off, txt in enumerate(window):
                    m1 = re.search(r"(\d{2})/(\d{2})/(\d{4})", txt)
                    if m1:
                        abspos = ws + idx_off
                        dist = abs(abspos - i)
                        if best is None or dist < best[0]:
                            best = (dist, m1.group(0))
                if best:
                    competencia_raw = best[1]

            # normaliza a competência para MM/YYYY (robusto)
            competencia = _normalize_competencia_to_mm_yyyy(competencia_raw)

            # grava o registro (se achou CNPJ)
            if cnpj_num:
                registros.append({
                    "CNPJ": formatar_cnpj(cnpj_num),
                    "Balancete_Protocolo": protocolo or "",
                    "Balancete_Competencia": competencia or ""
                })

            # avança o índice pra depois do bloco do participante (evita reprocessar)
            i = (j + 1) if j and (j + 1) > i else i + 1
            continue

        i += 1

    if not registros:
        return pd.DataFrame(columns=["CNPJ", "Balancete_Protocolo", "Balancete_Competencia"])

    df = pd.DataFrame(registros).drop_duplicates(subset="CNPJ", keep="first").reset_index(drop=True)
    return df


# --- Substitua também a função de PDF para usar o mesmo normalizador
def parse_protocolo_balancete_from_pdf(uploaded_pdf) -> pd.DataFrame:
    text = _read_text_from_pdf(uploaded_pdf)
    if not text:
        return pd.DataFrame(columns=["CNPJ", "Balancete_Protocolo", "Balancete_Competencia"])

    pattern_cnpj = re.compile(r"(\d{2}\.\d{3}\.\d{3}/\d{4}-\d{2})")
    pattern_proto = re.compile(r"(?:N[º°]\s*PROTOCOLO|PROTOCOLO)\D*([A-Z0-9\-]{4,60})", flags=re.I)
    registros = []

    for m in pattern_cnpj.finditer(text):
        cnpj_masked = m.group(1)
        start = m.start()
        window = text[max(0, start-400): start+500]

        # protocolo
        proto_m = pattern_proto.search(window)
        protocolo = proto_m.group(1).strip() if proto_m else ""

        # tentativa explícita de encontrar "Competência" no trecho
        comp_raw = None
        comp_match = re.search(r"(?:COMPETENCI[AÉE])\s*[:\-]?\s*([^\n\r]+)", window, flags=re.I)
        if comp_match:
            comp_raw = comp_match.group(1).strip()
        else:
            # procura mm/yyyy ou dd/mm/yyyy ou abreviação no trecho anterior
            m_comp = re.search(r"(\d{2}/\d{4}|\d{2}/\d{2}/\d{4}|[A-Za-zÀ-ÿ]{3,10}[/\s\-\.]*\d{2,4})", window)
            if m_comp:
                comp_raw = m_comp.group(1).strip()

        competencia = _normalize_competencia_to_mm_yyyy(comp_raw)

        cnpj_num = normaliza_cnpj(cnpj_masked)
        if cnpj_num:
            registros.append({
                "CNPJ": formatar_cnpj(cnpj_num),
                "Balancete_Protocolo": protocolo,
                "Balancete_Competencia": competencia or ""
            })

    if not registros:
        return pd.DataFrame(columns=["CNPJ", "Balancete_Protocolo", "Balancete_Competencia"])
    df = pd.DataFrame(registros)
    df = df.drop_duplicates(subset="CNPJ", keep="first").reset_index(drop=True)
    return df
