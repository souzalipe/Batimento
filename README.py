fundos_em_ambos_enriquecido = fundos_em_ambos_enriquecido.merge(
    df_balancete[["CNPJ", "Balancete_Protocolo", "Balancete_Competencia"]],
    on="CNPJ",
    how="left"
)
