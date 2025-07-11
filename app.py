import streamlit as st
import pandas as pd
import io
import os
from docx import Document
from src.components.merge_fields import *
from src.components.document_processing import process_docx_template
from src.components.mappings import load_default_mappings, load_uploaded_mappings

st.set_page_config(page_title="Brevkoder-automater", layout="wide")
st.title("Brevkoder-automater")


# Try to load default mapping at startup
DEFAULT_MAPPING_PATH = os.path.join("documents", "Liste over alle nøgler.xlsx")
default_mappings = load_default_mappings(DEFAULT_MAPPING_PATH)

# File upload for Excel
st.subheader("1. Upload Excel-fil med Titel/Nøgle-koblinger")

# Show info about default mapping below the upload field
if default_mappings:
    st.info(
        f"En standard-fil bruges allerede, men du kan uploade din egen, hvis du mener, at filen er forkert eller forældet."
    )

uploaded_file = st.file_uploader("Vælg en Excel-fil", type=["xlsx"])


mappings = None
if uploaded_file is not None:
    mappings, error = load_uploaded_mappings(uploaded_file)
    if error:
        st.error(error)
    else:
        st.write("Antal koblinger fundet:", len(mappings))
        with st.expander("Vis titel/nøgle-par (fra uploadet fil)", expanded=False):
            st.dataframe(
                pd.DataFrame(list(mappings.items()), columns=["Titel", "Nøgle"]),
                hide_index=True,
            )
elif default_mappings:
    st.write("Antal koblinger fundet:", len(default_mappings))
    with st.expander("Vis titel/nøgle-par (fra standardfil)", expanded=False):
        st.dataframe(
            pd.DataFrame(list(default_mappings.items()), columns=["Titel", "Nøgle"]),
            hide_index=True,
        )
    mappings = default_mappings
else:
    st.info("Upload venligst en Excel-fil med koblinger.")

if mappings:
    # Add file uploader for Word template and generate on upload
    st.subheader("2. Upload dit ukodede brev og generér kodet version")
    uploaded_docx = st.file_uploader(
        "Upload en .docx-fil som skabelon (dokumentet genereres automatisk ved upload)",
        type=["docx"],
    )

    # --- New: Text snippet processing ---
    st.subheader("2b. Indsæt tekststykke og generér kodet version")
    if "text_input" not in st.session_state:
        st.session_state["text_input"] = ""

    def update_text_output():
        pass  # No-op, just to trigger rerun

    text_input = st.text_area(
        "Indsæt et tekststykke her (f.eks. et afsnit fra et brev)",
        value=st.session_state["text_input"],
        height=150,
        placeholder="Indsæt tekst, der skal kodes med fletfelter...",
        key="text_input",
        on_change=update_text_output,
    )

    if st.session_state["text_input"].strip():
        from src.components.document_processing import (
            create_document_with_merge_fields,
        )

        _, debug_output = create_document_with_merge_fields(
            st.session_state["text_input"], mappings
        )
        st.markdown("**Kopierbar kodet tekst:**")
        st.text_area(
            "Resultat:",
            debug_output,
            height=150,
            key="output_text_snippet",
        )
    elif st.session_state["text_input"] != "":
        st.info("Indsæt venligst noget tekst for at generere kodet version.")
    # --- End new feature ---

    # --- Auto-load for testing if no upload ---
    if uploaded_docx is None:
        default_docx_path = os.path.join(
            "documents", "Ukodet dokument fra ønsket brevdesgin.docx"
        )
        if os.path.exists(default_docx_path):
            with open(default_docx_path, "rb") as f:
                uploaded_docx = io.BytesIO(f.read())
                uploaded_docx.name = "Ukodet dokument fra ønsket brevdesgin.docx"
    # --- End auto-load ---

    if uploaded_docx is not None:
        # Read the uploaded Word document and generate output immediately
        try:
            doc_template = Document(uploaded_docx)
            doc, debug_text = process_docx_template(doc_template, mappings)
            doc_io = io.BytesIO()
            doc.save(doc_io)
            doc_io.seek(0)
            st.subheader("3. Download det genererede dokument")
            st.success("Dokumentet er genereret!")
            st.download_button(
                label="Download Word-dokument",
                data=doc_io,
                file_name="dokument_med_fletfelter.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            )
        except Exception as e:
            st.error(f"Fejl ved behandling af Word-dokument: {str(e)}")
    else:
        st.info("Upload venligst en Word-skabelon.")

# Add instructions
with st.expander("Hjælp & Vejledning"):
    st.markdown(
        """
    ### Sådan bruger du denne app:
    
    1. **Upload en Excel-fil** der indeholder et ark med navnet 'query' og to kolonner:
       - "Titel": Tekststrenge, der optræder i dit brev
       - "Nøgle": Felt-navne, der skal bruges som Word-fletfelter
    
    2. **Upload en Word-skabelon**. Skabelonen skal indeholde de tekststrenge, der står i "Titel"-kolonnen.
       - Når du uploader skabelonen, genereres dokumentet automatisk.
    
    3. **Download** det færdige Word-dokument med fletfelter via knappen, der vises efter upload.
    
    Appen udskifter hver forekomst af tekst fra "Titel"-kolonnen med tilsvarende Word-fletfelter ud fra "Nøgle"-værdierne.
    """
    )
