

---

# Explicação da situação (Balancete_Protocolo)

## 1) Objetivo
Extrair corretamente o **Nº do Protocolo do Balancete** (formato `SCW\d+`) para cada **bloco de Participante** no relatório de Balancete e preencher a coluna `Balancete_Protocolo` no dataset final, casando por **CNPJ**.

## 2) Sintomas observados
- **Antes do patch**: a maioria dos protocolos extraídos ficava **deslocada** em relação ao valor correto (padrão “±1” no número do SCW).  
- **Depois do patch inicial (“só para frente”)**: a coluna `Balancete_Protocolo` passou a sair **vazia** para praticamente todos os fundos (exceto ETFs marcados como “Não possui”).

## 3) Causa-raiz
Existem **dois problemas distintos** que explicam os sintomas:

1. **Deslocamento (antes do patch)**  
   A busca do protocolo era feita **antes e depois** do “Participante”. Como os blocos são sequenciais, vasculhar **para trás** acabava pegando o protocolo do **bloco anterior**.

2. **Campos vazios (depois do patch)**  
   O patch “só para frente” foi correto na direção, mas a extração ficou **restritiva** demais:
   - Regex do rótulo **ancorada** (ex.: `^N[º°O]?\s*(DO\s*)?PROTOCOLO:?$`) falha quando o rótulo tem **texto residual** (ex.: “Protocolo do Balancete”, “Nº Protocolo :” etc.).
   - Suposição de que o valor está **sempre na célula imediatamente à direita**; em alguns layouts vem **duas colunas à direita** ou **na linha de baixo**.
   - Uso de `re.match` em vez de `re.search` para capturar o `SCW` quando há outros caracteres na célula.

## 4) Regras de negócio relevantes
- **Escopo do bloco**: tudo que está entre a célula que contém “**Participante**” do fundo atual **e** a próxima ocorrência de “Participante” (ou um limite de linhas) pertence ao **mesmo fundo**.
- **CNPJ** é capturado **nesse bloco** e o protocolo **deve** sair do mesmo bloco.
- **ETFs**: podem não ter protocolo de Balancete; nesses casos, `Balancete_Protocolo = "Não possui"`.
- **Competência do Balancete** (se já está correta no código) **não** deve ser alterada por este patch.

## 5) Solução (algoritmo robusto)
1. Após detectar o **“Participante”** e capturar o **CNPJ**, defina o **limite do bloco**: até a próxima linha que contenha “Participante” **ou** um teto de +60 linhas (o que ocorrer primeiro).  
2. **Tente primeiro via rótulo → valor**:  
   - Procure qualquer célula que **contenha** a palavra “PROTOCOLO” (não ancorar; não exigir “Nº”).  
   - Ao encontrar o rótulo, busque um `SCW\d{6,}`:
     - **na mesma linha**, nas próximas **1 a 3** células à direita;  
     - **na linha seguinte**, na **mesma coluna** do rótulo.  
3. **Fallback**: se não achar nessa heurística, capture o **primeiro `SCW\d{6,}`** que aparecer **em qualquer célula** dentro do bloco (antes do próximo “Participante”).  
4. **Valide** com regex `r'(SCW\d{6,})'` usando `re.search` (não `match`).  
5. **Pare** ao encontrar o primeiro protocolo válido para o bloco.

## 6) Bloco de código (para colar no lugar da busca do protocolo)
> **Observação**: este trecho supõe que você já tem `df_raw` como `DataFrame` (sem cabeçalho), `r` é a linha onde “Participante” foi encontrado, e `re` está importado.

```python
# --- Captura robusta do Nº do Protocolo (apenas para frente) ---
protocolo = None
upper = df_raw.shape[0]

# Define o limite do bloco: até o próximo 'Participante' ou +60 linhas
limite = min(r + 60, upper)
for rr2 in range(r + 1, limite):
    linha = df_raw.iloc[rr2].astype(str).tolist()
    if any('PARTICIPANTE' in str(x).strip().upper() for x in linha):
        limite = rr2  # fecha o bloco no início do próximo participante
        break

# 1) Heurística rótulo -> valor (mesma linha à direita; ou linha de baixo)
for rr in range(r, limite):
    row_vals = df_raw.iloc[rr].astype(str).tolist()
    for cc, cell in enumerate(row_vals):
        lab = str(cell).strip().upper()

        # Se a célula parece ser o rótulo de protocolo (qualquer variação contendo 'PROTOCOLO')
        if 'PROTOCOLO' in lab:
            # mesma linha: checa até +3 células à direita
            for k in range(cc + 1, min(cc + 4, len(row_vals))):
                m = re.search(r'(SCW\d{6,})', str(row_vals[k]))
                if m:
                    protocolo = m.group(1)
                    break
            if protocolo:
                break

            # linha de baixo: mesma coluna do rótulo
            if rr + 1 < upper:
                m = re.search(r'(SCW\d{6,})', str(df_raw.iat[rr + 1, cc]))
                if m:
                    protocolo = m.group(1)
                    break
    if protocolo:
        break

# 2) Fallback: primeiro SCW que aparecer no bloco
if not protocolo:
    for rr in range(r, limite):
        row_vals = df_raw.iloc[rr].astype(str).tolist()
        for cell in row_vals:
            m = re.search(r'(SCW\d{6,})', str(cell))
            if m:
                protocolo = m.group(1)
                break
        if protocolo:
            break
# --- fim ---
```

### Por que este bloco funciona
- **Direção correta** (só para frente): evita capturar o protocolo do bloco anterior.  
- **Rótulo flexível**: não depende de regex ancorada; cobre “Nº Protocolo”, “Protocolo do Balancete”, variações com “:” ou espaços.  
- **Posicionamento resiliente**: considera valor **à direita** e **logo abaixo** do rótulo.  
- **Plano B**: se o layout fugir do padrão, pega o **primeiro SCW** dentro do bloco (que é o desejado).

## 7) Erros comuns a evitar
- **`re.match`** para achar SCW: use `re.search` (o SCW pode não estar no início da célula).  
- Rótulo **ancorado** com `^...$`: layouts reais costumam ter texto extra.  
- Procurar **para trás** do “Participante”: causa o deslocamento “±1”.

## 8) Critérios de aceitação (sem depender de arquivos)
- Dado um bloco com “Participante …”, **CNPJ** e rótulo contendo “PROTOCOLO”, a função retorna um `SCW\d+` válido.  
- Quando existir mais de um `SCW` no bloco, retorna o **primeiro** após o rótulo ou, na falta de rótulo, o **primeiro do bloco**.  
- Para **ETFs**, manter `Balancete_Protocolo = "Não possui"` (ou `None`, conforme regra já existente).  
- O algoritmo **nunca** retorna o `SCW` do **bloco anterior**.

## 9) Testes (unitários simples, sem arquivos reais)
Monte `df_raw` mínimo com 2–3 blocos contendo:
- Variações de rótulo: “Nº Protocolo”, “Protocolo do Balancete:”.  
- Valor do SCW à direita, 2 células à direita e na linha abaixo.  
- Um `SCW` residual anterior (para garantir que **não** é capturado).  
- Um bloco **ETFs** sem protocolo (verificar “Não possui”).

Exemplo de assertiva (pseudocódigo):
```python
assert extrair_protocolo(df_raw, linha_participante_A) == "SCW202400123"
assert extrair_protocolo(df_raw, linha_participante_B) == "SCW202400456"
assert extrair_protocolo(df_raw, linha_participante_ETF) in (None, "Não possui")
```

---

Se você quiser, eu também te preparo uma versão com **logs de depuração** (ex.: `print`/`logger.debug` com a linha/coluna analisada) para facilitar ajuste fino quando aparecer um layout diferente. Quer?
