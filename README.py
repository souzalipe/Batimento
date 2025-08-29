def comparar_controle_fora_cadfi(cadfi_df, controle_df):
    """
    Retorna registros do Controle cujo CNPJ não aparece no CadFi (após padronização/duplicatas).
    Aplica também o filtro de status 'ativos' no Controle (exclui fundos encerrados).
    """
    if controle_df is None or controle_df.empty:
        return pd.DataFrame(columns=controle_df.columns if controle_df is not None else [])

    # 🔹 Aplica filtro de ativos no Controle
    controle_filtrado = filtrar_status_ativos(controle_df)

    return controle_filtrado[~controle_filtrado["CNPJ"].isin(set(cadfi_df["CNPJ"]))].copy()
