import xml.etree.ElementTree as ET
import pandas as pd

# -------------------------------------------------------------
# XML Routine Connection Analyzer for Rockwell L5X files
# -------------------------------------------------------------
# This script parses multiple Allen-Bradley Logix Designer (.L5X)
# project files, extracts wire connections between elements
# inside each Routine, and identifies "unconnected" IDs — i.e.,
# instruction references or operands that are not linked by any wire.
#
# The output is saved to an Excel file, listing all connections
# (Connected / Unconnected) for further review or debugging.
#
# Author: José Marques
# -------------------------------------------------------------


def processar_arquivo(caminho_arquivo):
    """
    Process a single .L5X file and collect wire and connection data.

    Args:
        caminho_arquivo (str): Path to the L5X file.
    """
    root = ler_arquivo_xml(caminho_arquivo)

    # Iterate over all routines in the controller
    for routine in root.findall(".//Routine"):
        wires, id_to_operand, plc_name, routine_name = process_routine(routine, root)

        # Identify IDs that are not connected by any wire
        unconnected = identify_unconnected(wires, id_to_operand)

        # Add "Unconnected" items to the global dataset
        append_unconnected_data(unconnected, plc_name, routine_name, id_to_operand)


def salvar_excel():
    """
    Generate a DataFrame and export the results to an Excel file.
    """
    df = pd.DataFrame(dados, columns=[
        "PLC",
        "Routine Name",
        "FromID",
        "FromOperand",
        "FromParam",
        "ToID",
        "ToOperand",
        "ToParam",
        "Status"
    ])

    # Sort the output for easier reading (Connected first, then Unconnected)
    df = df.sort_values(by=["Routine Name", "Status"], ascending=[True, True])

    # Export to Excel
    df.to_excel("resultado_wires_com_unconnected4.xlsx", index=False)

    print("✅ File 'resultado_wires_com_unconnected.xlsx' successfully generated!")


def ler_arquivo_xml(caminho_arquivo):
    """
    Read and parse an XML (.L5X) file.

    Args:
        caminho_arquivo (str): Path to the L5X file.

    Returns:
        xml.etree.ElementTree.Element: Root element of the XML tree.
    """
    try:
        tree = ET.parse(caminho_arquivo)
        root = tree.getroot()
        print('File successfully read:', root.tag)
        return root

    except ET.ParseError as e:
        print(f"XML parsing error: {e}")
    except FileNotFoundError:
        print(f"File not found: {caminho_arquivo}")
    except Exception as e:
        print(f"Unexpected error: {e}")


def process_routine(routine, root):
    """
    Process a single routine, mapping IDs and reading wire connections.

    Args:
        routine (Element): XML element for the Routine.
        root (Element): Root element of the entire L5X XML tree.

    Returns:
        tuple: (wires, id_to_operand, plc_name, routine_name)
    """
    routine_name = routine.attrib.get("Name", "")
    plc_name = root.find(".//Controller").attrib.get("Name", "Unknown_PLC")

    # Map all element IDs to their corresponding operand names
    id_to_operand = map_ID_operands(routine)

    # Extract all wire connections
    wires = []
    for wire in routine.findall(".//Wire"):
        from_id, to_id, from_param, to_param = read_wire(wire)
        wires.append((from_id, to_id, from_param, to_param))

        # Append this connected pair to the dataset
        append_connected_data(plc_name, routine_name, id_to_operand, from_id, from_param, to_id, to_param)

    return wires, id_to_operand, plc_name, routine_name


def identify_unconnected(wires, id_to_operand):
    """
    Identify IDs not linked by any wire in the routine.

    Args:
        wires (list): List of wire tuples (FromID, ToID, FromParam, ToParam)
        id_to_operand (dict): Mapping of IDs to operand names.

    Returns:
        set: IDs that are not connected.
    """
    connected_ids = set([f for f, _, _, _ in wires] + [t for _, t, _, _ in wires])
    all_ids = set(id_to_operand.keys())
    unconnected = all_ids - connected_ids
    return unconnected


def append_unconnected_data(unconnected, plc_name, routine_name, id_to_operand):
    """
    Append unconnected items to the global dataset.
    """
    for ref_id in unconnected:
        dados.append([
            plc_name,
            routine_name,
            ref_id,
            id_to_operand.get(ref_id, ""),
            "",
            "",
            "",
            "",
            "Unconnected"
        ])


def append_connected_data(plc_name, routine_name, id_to_operand, from_id, from_param, to_id, to_param):
    """
    Append connected wire data to the global dataset.
    """
    dados.append([
        plc_name,
        routine_name,
        from_id,
        id_to_operand.get(from_id, ""),
        from_param,
        to_id,
        id_to_operand.get(to_id, ""),
        to_param,
        "Connected"
    ])


def read_wire(wire):
    """
    Extract wire connection details from XML attributes.
    """
    from_id = wire.attrib.get("FromID")
    to_id = wire.attrib.get("ToID")
    from_param = wire.attrib.get("FromParam", "")
    to_param = wire.attrib.get("ToParam", "")
    return from_id, to_id, from_param, to_param


def map_ID_operands(routine):
    """
    Create a mapping of element IDs to their operands, names, or types.

    Args:
        routine (Element): XML element of the Routine.

    Returns:
        dict: {ID: Operand/Name/Type}
    """
    id_to_operand = {}
    for tag in (
        routine.findall(".//IRef")
        + routine.findall(".//ORef")
        + routine.findall(".//AddOnInstruction")
        + routine.findall(".//Function")
    ):
        ref_id = tag.attrib.get("ID")
        operand = tag.attrib.get("Operand", tag.attrib.get("Name", tag.attrib.get("Type", "")))
        if ref_id:
            id_to_operand[ref_id] = operand
    return id_to_operand


# -------------------------------------------------------------
# Main execution
# -------------------------------------------------------------
dados = []

# List of all L5X project files to be processed
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

# Process all files in sequence
for caminho_arquivo in arquivos_l5x:
    processar_arquivo(caminho_arquivo)

# Save the combined result
salvar_excel()
