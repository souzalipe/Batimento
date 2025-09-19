# supondo que seu arquivo de balancete seja 'balancete_file' (XLSX ou PDF)
# e que parse_protocolo_balancete / parse_protocolo_balancete_from_pdf
# estejam definidos conforme o seu CVM.py

# Ajuste o nome do uploaded_file abaixo ao rodar localmente
df_proto = parse_protocolo_balancete(balancete_file)  # se for xlsx
# df_proto = parse_protocolo_balancete_from_pdf(balancete_file)  # se for pdf

print("linhas extraídas pelo parser:", len(df_proto))
print(df_proto.head(20))
print("contagem CNPJ nulos:", df_proto["CNPJ"].isnull().sum() if "CNPJ" in df_proto.columns else "col CNPJ ausente")
print("contagem protocolos nulos:", df_proto["Balancete_Protocolo"].isnull().sum() if "Balancete_Protocolo" in df_proto.columns else "col Balancete_Protocolo ausente")
print("exemplos protocolos encontrados (value_counts):")
print(df_proto["Balancete_Protocolo"].value_counts(dropna=False).head(30))


