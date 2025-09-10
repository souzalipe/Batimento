# --- pega protocolo
protocolo = None
k = j if cnpj_fmt else i+1
while k < len(linhas) and protocolo is None:
    texto = linhas[k].strip()

    # Caso 1: protocolo na mesma célula (ex: "Nº Protocolo: 1234567")
    m_proto_inline = re.search(r"N[º°]?\s*PROTOCOLO[:\-]?\s*(\d{6,})", texto, flags=re.I)
    if m_proto_inline:
        protocolo = m_proto_inline.group(1)
        break

    # Caso 2: protocolo em célula ao lado (linha contém "Nº Protocolo" e a próxima célula é só o número)
    if texto.upper().startswith("Nº PROTOCOLO") or texto.upper().startswith("Nº DO PROTOCOLO"):
        if k + 1 < len(linhas):
            m_next = re.match(r"^\s*(\d{6,})\s*$", linhas[k+1])
            if m_next:
                protocolo = m_next.group(1)
                break

    # Caso 3: pular se for delimitador de bloco
    if "PROTOCOLO DE CONFIRMA" in texto.upper() or texto.upper().startswith("PARTICIPANTE"):
        break

    k += 1
