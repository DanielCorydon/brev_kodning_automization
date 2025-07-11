import os
import pandas as pd


def load_default_mappings(default_mapping_path):
    if os.path.exists(default_mapping_path):
        try:
            df_default = pd.read_excel(default_mapping_path, sheet_name="query")
            if "Titel" in df_default.columns and "Nøgle" in df_default.columns:
                return {row["Titel"]: row["Nøgle"] for _, row in df_default.iterrows()}
        except Exception:
            return None
    return None


def load_uploaded_mappings(uploaded_file):
    try:
        df = pd.read_excel(uploaded_file, sheet_name="query")
        if "Titel" not in df.columns or "Nøgle" not in df.columns:
            return (
                None,
                "Excel-filen skal indeholde kolonnerne 'Titel' og 'Nøgle' i arket 'query'.",
            )
        return {row["Titel"]: row["Nøgle"] for _, row in df.iterrows()}, None
    except Exception as e:
        return None, f"Fejl ved behandling af Excel-fil: {str(e)}"
