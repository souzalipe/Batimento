# --- Competência: procurar explicitamente a linha "Competência"
competencia = None
for idx, txt in enumerate(window):
    if "COMPETENCIA" in txt.upper() or "COMPETÊNCIA" in txt.upper():
        if idx + 1 < len(window):
            raw = window[idx+1].strip()

            # tenta dd/mm/yyyy
            m1 = re.match(r"(\d{2})/(\d{2})/(\d{4})", raw)
            if m1:
                competencia = f"{m1.group(2)}/{m1.group(3)}"
                break

            # tenta mm/yyyy
            m2 = re.match(r"(\d{2})/(\d{4})", raw)
            if m2:
                competencia = f"{m2.group(1)}/{m2.group(2)}"
                break

            # tenta jun/25 ou junho/25
            m3 = re.match(r"([A-Za-zÀ-ÿ]{3,10})[./\-\s]*(\d{2}|\d{4})", raw, flags=re.I)
            if m3:
                mes_txt = normaliza_texto(m3.group(1))
                mes_num = MESES_PT.get(mes_txt) or MESES_PT.get(mes_txt[:3])
                if mes_num:
                    ano_raw = m3.group(2)
                    ano = int(ano_raw) + 2000 if len(ano_raw) == 2 else int(ano_raw)
                    competencia = f"{mes_num:02d}/{ano}"
                    break

# fallback: se não achou após "Competência", usar a heurística antiga
if not competencia:
    best = None
    for idx_off, txt in enumerate(window):
        m1 = re.search(r"(\d{2})/(\d{2})/(\d{4})", txt)
        if m1:
            abspos = ws + idx_off
            dist = abs(abspos - i)
            if best is None or dist < best[0]:
                best = (dist, m1)
    if best:
        m = best[1]
        competencia = f"{m.group(2)}/{m.group(3)}"
