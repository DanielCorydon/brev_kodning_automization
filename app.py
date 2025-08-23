from docx import Document
import streamlit as st
import pandas as pd
import io
import os
import traceback
from loguru import logger
from src.components.mappings import load_default_mappings, load_uploaded_mappings
from src.components.agent import start_graph_llm, start_graph_llm_fake
from src.components.convert_text_fields import convert_document_fields


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


def save_docx_to_bytes(doc):
    """Save a Document object to BytesIO."""
    doc_io = io.BytesIO()
    doc.save(doc_io)
    doc_io.seek(0)
    return doc_io


# Add file uploader for Word template and generate on upload
st.subheader("2. Upload dit ukodede brev og generér kodet version")
uploaded_docx = st.file_uploader(
    "Upload en .docx-fil som skabelon (dokumentet genereres automatisk ved upload)",
    type=["docx"],
)

# --- Auto-load for testing if no upload ---
if uploaded_docx is None:
    # TEST_FILE_DOCS = "Ukodet dokument fra ønsket brevdesgin.docx"
    TEST_FILE_DOCS = "test_document_1.docx"
    default_docx_path = os.path.join("documents", TEST_FILE_DOCS)
    if os.path.exists(default_docx_path):
        with open(default_docx_path, "rb") as f:
            uploaded_docx = io.BytesIO(f.read())
            uploaded_docx.name = TEST_FILE_DOCS
# --- End auto-load ---

DEFAULT_SYSTEM_PROMPT = (
    'Du vil i teksten se et mønster hvor der står "if betingelse" (case insensitive), <en arbitrær mængde ord - lad os kalde dem MIDTERORD>, og så på et tidspunkt vil der stå ”<TEKST1>” Else ”<TEKST2>”. Altså 2 dobbelt anførselstegn med noget indeni, og så 2 dobbelt anførselstegn mere, med noget andet indeni. Her kalder vi dem TEKST1 og TEKST2, men du skal bruge det der står i teksten. '
    'Den første tekst-passage der passer på ovenstående mønster, skal erstattes med følgende: { IF "J" { MERGEFIELD <MIDTERORD>}" " TEKST1" " TEKST2" }'
)

st.subheader("Ekstra input til LLM (sendes som HumanMessage)")
extra_human_input = st.text_area(
    "Indtast ekstra tekst til LLM her:",
    value=DEFAULT_SYSTEM_PROMPT,
    height=200,
)

if uploaded_docx is not None:
    # Read the uploaded Word document and generate output immediately
    try:
        logger.debug("\n**-----------Processing uploaded document...-----------**\n\n")
        doc_bytes = uploaded_docx.read()
        output = start_graph_llm_fake(
            user_prompt=extra_human_input, document_bytes=doc_bytes
        )
        doc_bytes_out = output["document"][-1]
        doc_io = io.BytesIO(doc_bytes_out)
        doc = Document(doc_io)  # <-- Fix: load Document object
        doc = convert_document_fields(doc)  # <-- Pass Document, not BytesIO

        doc_io = save_docx_to_bytes(doc)
        st.subheader("4. Download det genererede dokument")
        st.success("Dokumentet er genereret!")
        st.download_button(
            label="Download Word-dokument",
            data=doc_io,
            file_name="dokument_med_fletfelter.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        )
    except Exception as e:
        error_type = type(e).__name__
        tb = traceback.format_exc()
        logger.debug(
            f"Fejl ved behandling af Word-dokument:\n\n"
            f"Type: {error_type}\n"
            f"Detaljer: {str(e)}\n\n"
            f"Traceback:\n{tb}"
        )
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
