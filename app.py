import streamlit as st
import pandas as pd
import io
import os
from src.components.mappings import load_default_mappings, load_uploaded_mappings
from src.components.process_document import (
    load_docx,
    save_docx_to_bytes,
    extract_colors_from_paragraph,
    process_paragraph_with_color_aware_replacements,
    insert_actual_mergefields,  # Added new import
    remove_comments_from_docx,  # Added new import
)
from src.components.find_fields import (
    transform_text_with_single_if_condition,
    transform_text_with_if_betingelse,
    replace_titles_with_mergefields,
)

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
    # Add text testing section
    st.subheader("1.5. Test transformation på tekst")
    st.write(
        "Du kan teste transformationen på en tekststreng her før du uploader et dokument."
    )

    # Default test text
    default_test_text = """Vi har gjort din Else til if betingelse Borger enlig ved ældrecheck berettigelse  ”og din ægtefælle/samlevers” likvide formue op på baggrund af If betingelse Borger enlig ved ældrecheck berettigelse ”din” Else ”jeres” årsopgørelse for Årstal forrige år fra Skattestyrelsen.

Din Else til if betingelse Borger enlig ved ældrecheck berettigelse”og din ægtefælle/samlevers” likvide formue har i Årstal indeværende år   været større end formuegrænsen. Formuegrænsen var Formuegrænse   kr. Det betyder, at du skal betale ældrechecken tilbage.
"""

    # Text input for testing
    test_text = st.text_area(
        "Indsæt tekst til test:",
        value="",
        height=150,
        placeholder='Eksempel: IF Betingelse KundeType "Privatkunde" ELSE "Erhvervskunde"',
    )

    # Function to transform text
    def transform_text(text, mappings):
        """
        Applies a series of transformations to the given text using the provided mappings.
        """
        transformed_text = transform_text_with_single_if_condition(text, mappings)
        transformed_text = transform_text_with_if_betingelse(transformed_text, mappings)
        transformed_text = replace_titles_with_mergefields(transformed_text, mappings)
        return transformed_text

    # Button to transform text
    if st.button("Transformér tekst"):
        if test_text.strip():
            # Apply transformations using the new function
            transformed_text = transform_text(test_text, mappings)

            st.subheader("Resultat:")
            st.text_area(
                "Transformeret tekst (kan kopieres):",
                value=transformed_text,
                height=150,
                key="result_text",
            )

            # Show if any changes were made
            if test_text != transformed_text:
                st.success("✅ Teksten blev transformeret!")
            else:
                st.info("ℹ️ Ingen 'IF Betingelse' mønstre fundet i teksten.")
        else:
            st.warning("Indsæt venligst noget tekst til transformation.")

    # Add file uploader for Word template and generate on upload
    st.subheader("2. Upload dit ukodede brev og generér kodet version")
    uploaded_docx = st.file_uploader(
        "Upload en .docx-fil som skabelon (dokumentet genereres automatisk ved upload)",
        type=["docx"],
    )

    if uploaded_docx is not None:
        # Read the uploaded Word document and generate output immediately
        try:
            doc_template = load_docx(uploaded_docx)

            # Remove comments from the document
            doc_template = remove_comments_from_docx(doc_template)

            # For each paragraph, first replace titles, then apply IF Betingelse transformation
            for para in doc_template.paragraphs:

                # Check for coloration in the paragraph
                text_colors, background_colors = extract_colors_from_paragraph(para)
                if text_colors or background_colors:
                    # Apply color-aware transformations for paragraphs with colors
                    para.text = process_paragraph_with_color_aware_replacements(
                        para, mappings
                    )

                    # Apply IF transformations on the already processed text
                    transformed_text = transform_text_with_single_if_condition(
                        para.text, mappings
                    )
                    transformed_text = transform_text_with_if_betingelse(
                        transformed_text, mappings
                    )

                    if para.text != transformed_text:
                        para.text = transformed_text

            # Convert all text-based MERGEFIELD syntax to actual Word merge fields
            for para in doc_template.paragraphs:
                insert_actual_mergefields(para)

            doc_io = save_docx_to_bytes(doc_template)
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
