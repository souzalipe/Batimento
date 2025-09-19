# --- captura CNPJ dentro do bloco (primeiro que aparecer)  (substituir o bloco atual) ---
cnpj_fmt = None
for rr in range(r, limite):
    for cell2 in df_raw.iloc[rr].astype(str).tolist():
        txt = str(cell2 or "")
        # 1) formato com pontuação
        m = re.search(r"(\d{2}\.\d{3}\.\d{3}/\d{4}-\d{2})", txt)
        if m:
            cnpj_fmt = m.group(1)
            break
        # 2) ou 14 dígitos seguidos (sem pontuação) -> formata
        digits = re.sub(r"\D", "", txt)
        if len(digits) == 14:
            cnpj_fmt = f"{digits[:2]}.{digits[2:5]}.{digits[5:8]}/{digits[8:12]}-{digits[12:]}"
            break
    if cnpj_fmt:
        break
cnpj_num = normaliza_cnpj(cnpj_fmt) if cnpj_fmt else None
