import streamlit as st
import tempfile
import os

# Importa as funcionalidades
from utils.parameters_extraction import gerar_planilhas_por_addon
from utils.addons_signatures import processar_addons
from utils.link_verification import main_link_verification


st.set_page_config(
    page_title="PLC Tools Web",
    page_icon="⚙️",
    layout="wide"
)

st.title("⚙️ PLC Toolkit Web")
st.write("Envie os arquivos `.L5X` e selecione as funcionalidades que deseja executar.")

# --- Upload de arquivos ---
arquivos = st.file_uploader(
    "Selecione um ou mais arquivos .L5X",
    type=["L5X", "xml"],
    accept_multiple_files=True
)

# --- Seleção de funcionalidades ---
st.subheader("🔧 Funcionalidades disponíveis:")
opcao_parameters_extraction = st.checkbox("Extração de parâmetros")
opcao_addons = st.checkbox("Gerar Add-ons Signatures (hash dos blocos)")
opcao_links = st.checkbox("Verificação de Links")


# --- Processamento ---
if arquivos and (opcao_addons or opcao_links):
    if st.button("▶️ Executar Processamento"):
        with st.spinner("Processando arquivos..."):
            with tempfile.TemporaryDirectory() as tmpdir:
                caminhos_locais = []
                for arquivo in arquivos:
                    caminho = os.path.join(tmpdir, arquivo.name)
                    with open(caminho, "wb") as f:
                        f.write(arquivo.read())
                    caminhos_locais.append(caminho)

                # Executa conforme as opções
                if opcao_parameters_extraction and False:
                    print('opção extração de parametros escolhida')
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

                if opcao_addons and False:
                    saida_addons = os.path.join(tmpdir, "addons_signatures.xlsx")
                    processar_addons(caminhos_locais, saida_addons)
                    with open(saida_addons, "rb") as f:
                        st.download_button(
                            "⬇️ Baixar resultado: Add-ons Signatures",
                            data=f,
                            file_name="addons_signatures.xlsx",
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                        )

                if opcao_links and False:
                    saida_links = os.path.join(tmpdir, "verificacao_links.xlsx")
                    main_link_verification(caminhos_locais, saida_links)
                    with open(saida_links, "rb") as f:
                        st.download_button(
                            "⬇️ Baixar resultado: Verificação de Links",
                            data=f,
                            file_name="verificacao_links.xlsx",
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                        )

        st.success("✅ Processamento concluído com sucesso!")
else:
    st.info("Envie arquivos e selecione ao menos uma funcionalidade para habilitar o processamento.")
