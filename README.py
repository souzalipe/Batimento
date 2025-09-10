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
