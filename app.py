import streamlit as st
import pandas as pd
import tempfile
import os
from utils.processing import gerar_planilhas_por_addon

st.set_page_config(page_title="Extrator de Add-Ons L5X", page_icon="⚙️", layout="centered")

st.title("⚙️ Extrator de Add-Ons de arquivos L5X")
st.write("Faça upload de um ou mais arquivos `.L5X` e gere uma planilha Excel com os dados extraídos.")

# Upload de múltiplos arquivos
arquivos = st.file_uploader("Selecione os arquivos L5X", type=["L5X", "xml"], accept_multiple_files=True)

if arquivos:
    st.info(f"{len(arquivos)} arquivo(s) carregado(s). Clique em 'Gerar Excel' para processar.")

    if st.button("Gerar Excel"):
        with st.spinner("Processando arquivos..."):
            # Cria diretório temporário para salvar uploads
            with tempfile.TemporaryDirectory() as tmpdir:
                caminhos_locais = []
                for arquivo in arquivos:
                    caminho_local = os.path.join(tmpdir, arquivo.name)
                    with open(caminho_local, "wb") as f:
                        f.write(arquivo.read())
                    caminhos_locais.append(caminho_local)

                caminho_saida = os.path.join(tmpdir, "dados_por_addon.xlsx")
                gerar_planilhas_por_addon(caminhos_locais, caminho_saida)

                # Retorna arquivo como download
                with open(caminho_saida, "rb") as f:
                    st.success("✅ Excel gerado com sucesso!")
                    st.download_button(
                        label="⬇️ Baixar Excel",
                        data=f,
                        file_name="dados_por_addon.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )
