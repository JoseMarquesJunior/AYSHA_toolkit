import xml.etree.ElementTree as ET
import pandas as pd
import os

def extrair_dados_tag(tag):
    """Extrai dados de uma tag XML e seus <DataValueMember>."""
    dados = {
        'Name': tag.get("Name"),
        'TagType': tag.get("TagType"),
        'DataType': tag.get("DataType"),
        'Constant': tag.get("Constant")
    }

    dados['DataValueMembers'] = [
        {"Name": d.get("Name"), "Value": d.get("Value")}
        for d in tag.findall(".//DataValueMember")
    ]
    return dados


def obter_lista_de_addons(caminho_arquivo):
    """Extrai todos os Add-On Instruction Definitions do arquivo XML."""
    try:
        tree = ET.parse(caminho_arquivo)
        root = tree.getroot()

        aoi_elements = root.findall('.//AddOnInstructionDefinition')
        lista = [aoi.get('Name') for aoi in aoi_elements if aoi.get('Name')]
        return lista
    except Exception as e:
        print(f"Erro ao obter Add-Ons de {caminho_arquivo}: {e}")
        return []


def ler_arquivo_xml(caminho_arquivo):
    """Retorna todas as tags com atributo DataType."""
    try:
        tree = ET.parse(caminho_arquivo)
        root = tree.getroot()
        return [extrair_dados_tag(tag) for tag in root.findall(".//Tag") if tag.get("DataType")]
    except Exception as e:
        print(f"Erro ao processar {caminho_arquivo}: {e}")
        return []


def gerar_planilhas_por_addon(arquivos, caminho_saida):
    """Processa arquivos XML e gera um Excel com uma aba por tipo de Add-On."""
    dados_por_addon = {}
    todos_addons = set()

    # --- Coletar todos os tipos de Add-Ons existentes ---
    for caminho in arquivos:
        addons_encontrados = obter_lista_de_addons(caminho)
        todos_addons.update(addons_encontrados)

    # --- Coletar tags e agrupar por tipo ---
    for caminho in arquivos:
        tags = ler_arquivo_xml(caminho)
        plc_name = os.path.basename(caminho).split('_')[-1].split('.')[0]

        for tag in tags:
            tipo = tag['DataType']
            if tipo in todos_addons:
                tag_data = {'PLC': plc_name, 'AddOn': tipo, 'Name': tag['Name']}
                tag_data.update({m['Name']: m['Value'] for m in tag['DataValueMembers']})
                dados_por_addon.setdefault(tipo, []).append(tag_data)

    # --- Criar Excel com abas ---
    os.makedirs(os.path.dirname(caminho_saida), exist_ok=True)
    with pd.ExcelWriter(caminho_saida) as writer:
        for tipo, registros in sorted(dados_por_addon.items()):
            df = pd.DataFrame(registros)
            df.to_excel(writer, sheet_name=tipo[:31], index=False)
