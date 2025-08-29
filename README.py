cadfi_filtrado = filtrar_cadfi(cadfi_raw)

# Aplica filtro de ativos também no Controle
controle_prep = carregar_controle(controle_raw)
controle_prep = filtrar_status_ativos(controle_prep)
