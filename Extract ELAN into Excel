##Creado con prompts en ChatGPT 5.4

from pathlib import Path
import xml.etree.ElementTree as ET
import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox


# Columnas estándar que tendrá el Excel
TARGET_TIERS = ["GLOSA", "NO MANUAL", "CLASIFICADORES", "TRADUCCION", "GLOSA_IA"]

# Equivalencias de nombres de tiers
TIER_EQUIVALENCES = {
    "GLOSA": ["GLOSA"],
    "NO MANUAL": ["NO MANUAL", "No manuales"],
    "CLASIFICADORES": ["CLASIFICADORES", "Clasificadores"],
    "TRADUCCION": ["TRADUCCION", "Traducción"],
    "GLOSA_IA": ["GLOSA_IA"]
}


def select_folder() -> Path | None:
    """Permite seleccionar la carpeta madre."""
    root = tk.Tk()
    root.withdraw()
    root.attributes("-topmost", True)

    folder_selected = filedialog.askdirectory(title="Selecciona la carpeta madre")
    root.destroy()

    if not folder_selected:
        return None

    return Path(folder_selected)


def clean_text(text: str) -> str:
    """Limpia espacios extra y saltos de línea."""
    if text is None:
        return ""
    return " ".join(text.split())


def normalize_tier_name(tier_name: str) -> str | None:
    """
    Devuelve el nombre estándar del tier si el nombre recibido
    coincide con alguna equivalencia.
    """
    for standard_name, variants in TIER_EQUIVALENCES.items():
        if tier_name in variants:
            return standard_name
    return None


def extract_tier_annotations(eaf_file: Path, parent_folder_name: str) -> dict:
    """
    Extrae las anotaciones de los tiers de interés desde un archivo .eaf.
    Devuelve una fila lista para el Excel.
    """
    row = {
        "Carpeta mayor": parent_folder_name,
        "Archivo": eaf_file.name
    }

    for tier_name in TARGET_TIERS:
        row[tier_name] = ""

    try:
        tree = ET.parse(eaf_file)
        root = tree.getroot()
    except ET.ParseError as e:
        print(f"Error al parsear {eaf_file}: {e}")
        return row
    except Exception as e:
        print(f"Error al leer {eaf_file}: {e}")
        return row

    for tier in root.findall("TIER"):
        tier_id = tier.attrib.get("TIER_ID", "").strip()
        normalized_name = normalize_tier_name(tier_id)

        if normalized_name is None:
            continue

        annotations = []

        for annotation in tier.findall("ANNOTATION"):
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

        joined_annotations = " | ".join(annotations)

        # Si ya había contenido en esa columna, lo concatenamos
        if joined_annotations:
            if row[normalized_name]:
                row[normalized_name] += " | " + joined_annotations
            else:
                row[normalized_name] = joined_annotations

    return row


def process_parent_folder(parent_folder: Path) -> pd.DataFrame:
    """
    Recorre las subcarpetas de la carpeta madre y procesa todos los .eaf.
    Genera una fila por archivo.
    """
    rows = []

    subfolders = [f for f in parent_folder.iterdir() if f.is_dir()]

    if not subfolders:
        raise FileNotFoundError("No se encontraron subcarpetas dentro de la carpeta madre.")

    for subfolder in sorted(subfolders):
        eaf_files = sorted(subfolder.glob("*.eaf"))

        if not eaf_files:
            print(f"No se encontraron archivos .eaf en: {subfolder.name}")
            continue

        for eaf_file in eaf_files:
            print(f"Procesando: {subfolder.name} -> {eaf_file.name}")
            row = extract_tier_annotations(eaf_file, subfolder.name)
            rows.append(row)

    if not rows:
        raise FileNotFoundError("No se encontraron archivos .eaf en las subcarpetas de la carpeta madre.")

    columns = ["Carpeta mayor", "Archivo"] + TARGET_TIERS
    df = pd.DataFrame(rows, columns=columns)
    return df


def save_to_excel(df: pd.DataFrame, parent_folder: Path) -> Path:
    """Guarda el resultado en un Excel dentro de la carpeta madre."""
    output_file = parent_folder / "extraccion_tiers_eaf.xlsx"

    with pd.ExcelWriter(output_file, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="ELAN")

        worksheet = writer.sheets["ELAN"]

        for column_cells in worksheet.columns:
            max_length = 0
            column_letter = column_cells[0].column_letter

            for cell in column_cells:
                cell_value = str(cell.value) if cell.value is not None else ""
                max_length = max(max_length, len(cell_value))

            worksheet.column_dimensions[column_letter].width = min(max_length + 2, 60)

    return output_file


def main():
    parent_folder = select_folder()

    if parent_folder is None:
        print("No se seleccionó ninguna carpeta.")
        return

    try:
        df = process_parent_folder(parent_folder)
        output_file = save_to_excel(df, parent_folder)

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
