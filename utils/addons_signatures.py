import xml.etree.ElementTree as ET
import pandas as pd
import hashlib

def ler_arquivo_xml(caminho_arquivo):
    """
    Função para ler e analisar o arquivo XML e extrair as tags com DataType 'AI_TYPE5'.

    Args:
        caminho_arquivo (str): Caminho para o arquivo XML.

    Returns:
        list: Lista de dicionários com os dados extraídos das tags relevantes.
    """
    try:
        tree = ET.parse(caminho_arquivo)
        root = tree.getroot()
        print('Arquivo lido com sucesso:', root.tag)  
        return root
    
    except ET.ParseError as e:
        print(f"Erro ao analisar o arquivo XML: {e}")
    except FileNotFoundError:
        print(f"Arquivo não encontrado: {caminho_arquivo}")
    except Exception as e:
        print(f"Erro inesperado: {e}")

def calcular_hash(aoi, target_name):
    # Nome do bloco
    aoi_name = aoi.attrib.get('Name', 'SemNome')

    # Serializa o XML da subárvore em string
    xml_bytes = ET.tostring(aoi, encoding="utf-8")
    xml_str = xml_bytes.decode("utf-8")   # <-- converte bytes em string

    # Divide a partir da primeira ocorrência de <Parameters>
    parte_parameters = xml_str.split("<Parameters>", 1)[-1]

    # Reconstroi incluindo a tag <Parameters>
    novo_xml = "<Parameters>" + parte_parameters

    # Calcula hash MD5
    hash_calculado = hashlib.md5(novo_xml.encode("utf-8")).hexdigest()
    hash_numerico = int(hash_calculado, 16)

    if aoi_name == 'AI_TYPE5':
        nome_arquivo = f"{target_name}_{aoi_name}.txt"
        with open(nome_arquivo, "w", encoding="utf-8") as f:
            f.write(novo_xml)
            

    return hash_numerico, aoi_name

def processar_arquivo(caminho_arquivo):
    root = ler_arquivo_xml(caminho_arquivo)
    if root:
        # Pegar o atributo TargetName, que representa o nome do PLC
        target_name = root.attrib.get("TargetName", "SemTargetName")
        print(target_name)
        
        # Buscar todos os AddOnInstructionDefinition
        aoi_elements = root.findall('.//AddOnInstructionDefinition')

        
        for aoi in aoi_elements:   
            hash_numerico, aoi_name = calcular_hash(aoi, target_name) 
            dados.append([target_name, aoi_name, hash_numerico])

def salvar_excel(dados):
    # Criar DataFrame
    dados_listados = pd.DataFrame(dados, columns=['PLC', 'Add-On', 'HASH'])

    dados_pivotados = dados_listados.pivot(
            index="Add-On", 
            columns="PLC", 
            values="HASH"
            ).reset_index()
    

    # Salvar em Excel
    print("Planilha criada com sucesso!")

    with pd.ExcelWriter("Controle de blocos/bkp_TS_20251011_Hull_20251007/Controle de blocos.xlsx", engine="openpyxl") as writer:
        dados_pivotados.to_excel(writer, sheet_name="bkp_TS_20251011_Hull_20251007", index=False)
        dados_listados.to_excel(writer, sheet_name="Lista", index=False)
    
    print("Planilha criada com sucesso!")

dados = []
hashes = {}

# Lista de todos os arquivos L5X
arquivos_l5x = [
    "backups_PLC/Topsides/20251011/P80_TOPSIDE_PCS01.L5X",
    "backups_PLC/Topsides/20251011/P80_TOPSIDE_PCS02.L5X",
    "backups_PLC/Topsides/20251011/P80_TOPSIDE_PCS03.L5X",
    "backups_PLC/Topsides/20251011/P80_TOPSIDE_PSD01.L5X",
    "backups_PLC/Topsides/20251011/P80_TOPSIDE_PSD02.L5X",
    "backups_PLC/Topsides/20251011/P80_TOPSIDE_PSD03.L5X",
    "backups_PLC/Topsides/20251011/P80_TOPSIDE_FGS01.L5X",
    "backups_PLC/Topsides/20251011/P80_TOPSIDE_FGS02.L5X",
    "backups_PLC/Hull/20251007/P80_HULL_HCS01.L5X",
    "backups_PLC/Hull/20251007/P80_HULL_HCS02.L5X",
    "backups_PLC/Hull/20251007/P80_HULL_HSD01.L5X",
    "backups_PLC/Hull/20251007/P80_HULLSIDE_HFGS01.L5X",
    "backups_PLC/Hull/20251007/P80_HULLSIDE_HFGS02.L5X"
]

# Loop para processar todos os arquivos
for caminho in arquivos_l5x:
    processar_arquivo(caminho)

salvar_excel(dados)





