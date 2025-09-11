# procurar competência ANTES do participante
competencia = None
for rr in range(max(0, r-15), r+1):
    row_vals = df_raw.iloc[rr].tolist()
    for cc, cell in enumerate(row_vals):
        if isinstance(cell, str) and cell.strip().upper().startswith("COMPETÊNCIA"):
            if cc + 1 < len(row_vals):
                competencia_raw = row_vals[cc+1]
                # se for data do Excel, converte
                try:
                    comp_date = pd.to_datetime(competencia_raw, errors="coerce")
                    if pd.notna(comp_date):
                        competencia = f"{comp_date.month:02d}/{comp_date.year}"
                    else:
                        competencia = str(competencia_raw).strip()
                except Exception:
                    competencia = str(competencia_raw).strip()
            break
    if competencia:
        break
