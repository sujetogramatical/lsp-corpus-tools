from pathlib import Path
import xml.etree.ElementTree as ET
import pandas as pd
import tkinter as tk
from tkinter import filedialog
import unicodedata


# Lista de negación base
NEGATION_WORDS = [
    "NO", "NADA", "NADIE", "NINGUN", "NINGUNO", "NINGUNA",
    "NUNCA", "JAMAS", "TAMPOCO", "NI", "FALTAR"
]


def normalize_text(text: str) -> str:
    """Normaliza texto: mayúsculas + sin tildes"""
    text = text.upper()
    text = unicodedata.normalize('NFD', text)
    text = ''.join(c for c in text if unicodedata.category(c) != 'Mn')
    return text.strip()


def select_folder() -> Path | None:
    root = tk.Tk()
    root.withdraw()
    root.attributes("-topmost", True)

    folder_selected = filedialog.askdirectory(title="Selecciona la carpeta con archivos EAF")
    root.destroy()

    if not folder_selected:
        return None

    return Path(folder_selected)


def extract_glosa_annotations(eaf_file: Path) -> list[str]:
    """Extrae anotaciones del tier GLOSA"""
    try:
        tree = ET.parse(eaf_file)
        root = tree.getroot()
    except:
        return []

    for tier in root.findall("TIER"):
        if tier.attrib.get("TIER_ID") == "GLOSA":
            annotations = []

            for annotation in tier.findall("ANNOTATION"):
                alignable = annotation.find("ALIGNABLE_ANNOTATION")
                ref_ann = annotation.find("REF_ANNOTATION")

                if alignable is not None:
                    val = alignable.find("ANNOTATION_VALUE")
                    if val is not None and val.text:
                        annotations.append(val.text.strip())

                elif ref_ann is not None:
                    val = ref_ann.find("ANNOTATION_VALUE")
                    if val is not None and val.text:
                        annotations.append(val.text.strip())

            return annotations

    return []


def analyze_negation(folder_path: Path):
    resumen_rows = []
    ocurrencias_rows = []

    eaf_files = sorted(folder_path.glob("*.eaf"))

    for eaf_file in eaf_files:
        print(f"Procesando: {eaf_file.name}")

        glosas = extract_glosa_annotations(eaf_file)
        glosas_norm = [normalize_text(g) for g in glosas]

        conteo = {w: 0 for w in NEGATION_WORDS}
        conteo["NO-VERBO"] = 0

        coincidencias = []

        for i, (original, norm) in enumerate(zip(glosas, glosas_norm)):

            match = None

            # Caso 1: coincidencia exacta
            if norm in NEGATION_WORDS:
                match = norm
                conteo[norm] += 1

            # Caso 2: NO-VERBO
            elif norm.startswith("NO-"):
                match = "NO-VERBO"
                conteo["NO-VERBO"] += 1

            if match:
                coincidencias.append(original)

                contexto_prev = glosas[i-1] if i > 0 else ""
                contexto_next = glosas[i+1] if i < len(glosas)-1 else ""

                ocurrencias_rows.append({
                    "Archivo": eaf_file.name,
                    "Glosa original": original,
                    "Forma normalizada": norm,
                    "Tipo": match,
                    "Posición": i + 1,
                    "Previo": contexto_prev,
                    "Siguiente": contexto_next
                })

        resumen_row = {
            "Archivo": eaf_file.name,
            "Total": sum(conteo.values()),
            "Coincidencias": " | ".join(coincidencias)
        }

        resumen_row.update(conteo)
        resumen_rows.append(resumen_row)

    df_resumen = pd.DataFrame(resumen_rows)
    df_ocurrencias = pd.DataFrame(ocurrencias_rows)

    return df_resumen, df_ocurrencias


def save_excel(df_resumen, df_ocurrencias, folder_path):
    output_file = folder_path / "negacion_resultados.xlsx"

    with pd.ExcelWriter(output_file, engine="openpyxl") as writer:
        df_resumen.to_excel(writer, index=False, sheet_name="RESUMEN")
        df_ocurrencias.to_excel(writer, index=False, sheet_name="OCURRENCIAS")

    return output_file


def main():
    folder = select_folder()

    if folder is None:
        print("No se seleccionó carpeta")
        return

    df_resumen, df_ocurrencias = analyze_negation(folder)
    output = save_excel(df_resumen, df_ocurrencias, folder)

    print(f"\nExcel generado en: {output}")


if __name__ == "__main__":
    main()