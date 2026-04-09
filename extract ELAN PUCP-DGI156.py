from pathlib import Path
import xml.etree.ElementTree as ET
import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox


# Columnas principales del Excel
TARGET_TIERS = ["GLOSA", "CLASIFICADORES", "NO MANUAL", "DESCRIBIR"]

# Equivalencias de nombres de tiers
TIER_EQUIVALENCES = {
    "GLOSA": ["GLOSA"],
    "CLASIFICADORES": ["CLASIFICADORES", "Clasificadores"],
    "NO MANUAL": ["NO MANUAL", "No manuales"],
    "DESCRIBIR": ["DESCRIBIR", "Describir"]
}


def select_folder() -> Path | None:
    root = tk.Tk()
    root.withdraw()
    root.attributes("-topmost", True)

    folder_selected = filedialog.askdirectory(title="Selecciona la carpeta con archivos EAF")
    root.destroy()

    if not folder_selected:
        return None

    return Path(folder_selected)


def clean_text(text: str) -> str:
    if text is None:
        return ""
    return " ".join(text.split())


def normalize_tier_name(tier_name: str) -> str | None:
    for standard_name, variants in TIER_EQUIVALENCES.items():
        if tier_name in variants:
            return standard_name
    return None


def extract_annotation_values(tier_element) -> list[str]:
    annotations = []

    for annotation in tier_element.findall("ANNOTATION"):
        alignable = annotation.find("ALIGNABLE_ANNOTATION")
        ref_ann = annotation.find("REF_ANNOTATION")

        if alignable is not None:
            ann_value = alignable.find("ANNOTATION_VALUE")
            if ann_value is not None and ann_value.text:
                annotations.append(clean_text(ann_value.text))

        elif ref_ann is not None:
            ann_value = ref_ann.find("ANNOTATION_VALUE")
            if ann_value is not None and ann_value.text:
                annotations.append(clean_text(ann_value.text))

    return annotations


def extract_tier_annotations(eaf_file: Path) -> dict:
    row = {"Archivo": eaf_file.name}

    # Inicializamos todo como inexistente
    for tier_name in TARGET_TIERS:
        row[tier_name] = "Tier inexistente"

    row["OTROS"] = ""

    try:
        tree = ET.parse(eaf_file)
        root = tree.getroot()
    except Exception as e:
        print(f"Error al leer {eaf_file}: {e}")
        row["OTROS"] = "Error al leer el archivo"
        return row

    other_tiers = []
    found_tiers = set()

    for tier in root.findall("TIER"):
        tier_id = tier.attrib.get("TIER_ID", "").strip()
        normalized_name = normalize_tier_name(tier_id)

        if normalized_name is not None:
            found_tiers.add(normalized_name)

            annotations = extract_annotation_values(tier)

            if annotations:
                row[normalized_name] = " | ".join(annotations)
            else:
                row[normalized_name] = "Tier sin datos"

        else:
            other_tiers.append(f"Tiene un tier {tier_id}")

    # Aseguramos que los tiers encontrados pero vacíos estén bien marcados
    for tier_name in TARGET_TIERS:
        if tier_name in found_tiers and row[tier_name] == "Tier inexistente":
            row[tier_name] = "Tier sin datos"

    row["OTROS"] = " | ".join(other_tiers) if other_tiers else ""

    return row


def process_folder(folder_path: Path) -> pd.DataFrame:
    eaf_files = sorted(folder_path.glob("*.eaf"))

    if not eaf_files:
        raise FileNotFoundError("No se encontraron archivos .eaf en la carpeta seleccionada.")

    rows = []

    for eaf_file in eaf_files:
        print(f"Procesando: {eaf_file.name}")
        row = extract_tier_annotations(eaf_file)
        rows.append(row)

    columns = ["Archivo"] + TARGET_TIERS + ["OTROS"]
    df = pd.DataFrame(rows, columns=columns)
    return df


def save_to_excel(df: pd.DataFrame, folder_path: Path) -> Path:
    output_file = folder_path / "PUCP-DGI156_extraccion_tiers_eaf.xlsx"

    with pd.ExcelWriter(output_file, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="ELAN")

        worksheet = writer.sheets["ELAN"]

        for column_cells in worksheet.columns:
            max_length = 0
            column_letter = column_cells[0].column_letter

            for cell in column_cells:
                cell_value = str(cell.value) if cell.value is not None else ""
                max_length = max(max_length, len(cell_value))

            worksheet.column_dimensions[column_letter].width = min(max_length + 2, 80)

    return output_file


def main():
    folder_path = select_folder()

    if folder_path is None:
        print("No se seleccionó ninguna carpeta.")
        return

    try:
        df = process_folder(folder_path)
        output_file = save_to_excel(df, folder_path)

        print("\nProceso completado.")
        print(f"Archivo Excel guardado en: {output_file}")

        root = tk.Tk()
        root.withdraw()
        messagebox.showinfo(
            "Proceso completado",
            f"El archivo Excel se guardó correctamente en:\n{output_file}"
        )
        root.destroy()

    except Exception as e:
        print(f"Ocurrió un error: {e}")

        root = tk.Tk()
        root.withdraw()
        messagebox.showerror("Error", f"Ocurrió un error:\n{e}")
        root.destroy()


if __name__ == "__main__":
    main()