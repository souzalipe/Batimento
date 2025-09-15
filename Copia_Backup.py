import io
import zipfile
import pandas as pd
import streamlit as st

def to_excel_bytes(df: pd.DataFrame, sheet_name="Resultado") -> bytes:
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
        df.to_excel(writer, index=False, sheet_name=sheet_name)
    return output.getvalue()

# --- no seu fluxo real, esses DataFrames já existem ---
# exemplo didático (troque pelos seus df_resultado1,2,3)
df_resultado1 = pd.DataFrame({"CNPJ": ["11.111.111/1111-11"], "Status": ["OK"]})
df_resultado2 = pd.DataFrame({"CNPJ": ["22.222.222/2222-22"], "Controle": ["Faltando"]})
df_resultado3 = pd.DataFrame({"CNPJ": ["33.333.333/3333-33"], "Cadfi": ["Excedente"]})

# --- cria um ZIP em memória ---
zip_buffer = io.BytesIO()
with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zipf:
    zipf.writestr("batimento_resultado1.xlsx", to_excel_bytes(df_resultado1))
    zipf.writestr("batimento_resultado2.xlsx", to_excel_bytes(df_resultado2))
    zipf.writestr("batimento_resultado3.xlsx", to_excel_bytes(df_resultado3))

# reposiciona ponteiro do buffer
zip_buffer.seek(0)

# --- botão único de download ---
st.download_button(
    label="⬇️ Baixar todos os resultados (ZIP)",
    data=zip_buffer,
    file_name="batimento_fundos.zip",
    mime="application/zip"
)

k
