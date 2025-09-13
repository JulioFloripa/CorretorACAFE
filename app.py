import streamlit as st

st.set_page_config(page_title="Gerador de Boletins - Fleming", layout="centered")
st.title("📊 Gerador de Boletins - Colégio Fleming")

st.markdown("""
Faça upload da planilha do simulado no formato `.xlsx` com as abas:
- `RESPOSTAS`
- `GABARITO`

O sistema irá gerar automaticamente os boletins individuais em PDF para cada flemer.
""")

# Upload do arquivo
arquivo = st.file_uploader("📎 Faça upload da planilha do simulado", type=["xlsx"])

if arquivo:
    st.success("✅ Planilha recebida com sucesso!")
    st.info("Aqui vamos inserir a lógica para gerar os PDFs e gráficos.")
    # TODO: inserir aqui a lógica de geração de boletins
