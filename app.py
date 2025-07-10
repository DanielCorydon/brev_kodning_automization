import streamlit as st
import pandas as pd
from docx import Document
from docx.oxml import OxmlElement, parse_xml
from docx.oxml.ns import qn
from docx.enum.text import WD_PARAGRAPH_ALIGNMENT
import re
import io
import os

st.set_page_config(page_title="Brevkoder-automater", layout="wide")
st.title("Brevkoder-automater")

# Define default test template text
DEFAULT_TEMPLATE_TEXT = """IF Betingelse Ubegrænset fuldmagt
"Du modtager dette brev på vegne af [Fuldmagtsgivers navn], som du er fuldmagtshaver eller værge for."

Du skal betale ældrecheck for Årstal indeværende år   tilbage
Vi skriver til dig, fordi skriv hvem der har givet os oplysninger har givet os nye oplysninger om If betingelse Borger enlig ved ældrecheck berettigelse"din" Else "jeres" likvide formue efter den seneste opgørelse i brevet Skriv titel på brev og dato for udsendelse. De nye oplysninger ændrer ikke den seneste opgørelse af If betingelse Borger enlig ved ældrecheck berettigelse  "din" Else "jeres" likvide formue til ældrecheck Årstal indeværende år  .

Vores opgørelse viser, at du uberettiget har fået udbetalt ældrecheck og fortsat skal betale Kontrol af ældrecheck krav kr. efter skat  tilbage.

Du kan se opgørelsen af If betingelse Borger enlig ved ældrecheck berettigelse "din" Else "jeres" likvide formue længere nede i brevet. 

Det skal du gøre 
Du kan vælge at: 
•	betale med betalingskort eller MobilePay på www.borger.dk/betal. Her kan du også oprette en tilbagebetalingsordning 
•	betale beløbet via det indbetalingskort, du får indenfor 3 uger. 
Vær opmærksom på, at det er helt frivilligt, om du vil benytte dig af ovennævnte betalingsmuligheder.

Der lægges ikke renter på den pension, du skal betale tilbage.

Hvis du ikke gør noget 
Hvis du ikke betaler/opretter afdragsordning på borger.dk eller betaler via indbetalingskortet, beregner vi, hvordan du skal betale beløbet tilbage. Her ser vi på dine økonomiske forhold, bl.a. indkomster, udgifter og formue. Det vil du få et brev om.

(S2)"Har du spørgsmål?
Hvis du har spørgsmål eller er uenig i vores afgørelse, er du velkommen til at ringe til os på telefon 70 12 80 61.

Du kan læse mere om folkepension på www.borger.dk/folkepension."


Venlig hilsen

Udbetaling Danmark
Pension




Begrundelse for afgørelsen
Vi har gjort din Else til if betingelse Borger enlig ved ældrecheck berettigelse  "og din ægtefælle/ samlevers" likvide formue op på baggrund af If betingelse Borger enlig ved ældrecheck berettigelse "din" Else "jeres" årsopgørelse for Årstal forrige   år fra Skattestyrelsen.

Din Else til if betingelse Borger enlig ved ældrecheck berettigelse"og din ægtefælle/samlevers" likvide formue har i Årstal indeværende år   været større end formuegrænsen. Formuegrænsen var Formuegrænse   kr. Det betyder, at du skal betale ældrechecken tilbage.

Du har oplyst, at Skriv hvilke oplysninger vi har fået fra borger.

Skriv hvorfor vi ikke har ændret i den tidligere afgørelse

Når formuen er højere end formuegrænsen, har man ikke ret til ældrecheck. Formuegrænsen var Formuegrænse   kr. i Årstal indeværende år  . Opgørelsen af din formue viser, at din Else til if betingelse Borger enlig ved ældrecheck berettigelse "og din ægtefælle/samlevers" likvide formue udgjorde Faktisk formue samlet pensionist og samlever. Du havde derfor ikke ret til ældrecheck for årstal indeværende år  .

Når man får ældrecheck udbetalt, skal man betale beløbet tilbage, hvis man ved eller bør vide, at man ikke har ret til pengene, når man modtager dem. Vi har vurderet, at du ikke havde grund til at tro, at du kunne få ældrecheck udbetalt, og kræver derfor beløbet tilbage

Vi lægger ved vores vurdering vægt på, at det fremgik i afgørelsen om ældrecheck, at du var berettiget til ældrecheck, fordi din formue pr. 1. januar Årstal indeværende år   var under formuegrænsen på Formuegrænse   kr.

Vi lægger desuden vægt på, at det også af afgørelsen om ældrecheck fremgik, at vi ved opgørelsen af din likvide formue havde hentet If betingelse Borger enlig ved ældrecheck berettigelse "dine" Else "din og din samlever/ægtefælles" formueoplysninger fra seneste årsopgørelse fra Skattestyrelsen eller If betingelse Borger enlig ved ældrecheck berettigelse"dine" Else "din og din samlevers/ægtefælles" egne oplysninger. Den samlede likvide formue udgjorde Bevillingsformue samlet pensionist og samlever.

Vi oplyste dig om, at du skulle give os besked, hvis din formue pr. 1. januar Årstal indeværende år   oversteg formuegrænsen. Du har ikke givet os besked herom.

Det er derfor vores samlede vurdering, at du vidste eller burde vide, da du fik udbetalt ældrecheck for Årstal indeværende år  , at den likvide formue var højere end formuegrænsen.



Sådan har vi gjort den likvide formue op
Her kan du se, hvordan vi har gjort din Else til if betingelse Borger enlig ved ældrecheck berettigelse  "og din ægtefælle/samlevers" likvide formue op. Vi har hentet oplysningerne fra årsopgørelsen fra Skattestyrelsen for Årstal forrige år  ,og de oplysninger om din likvide formue, If betingelse Borger enlig ved ældrecheck berettigelse If Manuel betingelse Borger beregnes som enlig "du" Else "I" selv har givet os.

IF Betingelse Formue oplyst af borger 
"*Beløbene er oplyst af dig Else til if betingelse Borger enlig ved ældrecheck berettigelse  "og/eller din ægtefælle/samlevers"."

B17A – FOP Formue opgørelse – Ældrecheck besvigelseskontrol

  
(M2) "Vi kan modregne i overskydende skat
Ifølge databeskyttelsesforordningen og databeskyttelsesloven skal vi orientere dig om, at det beløb, du skylder os, vil blive registreret hos Gældsstyrelsen, indtil det skyldige beløb er betalt. Oplysningerne i registret kan bruges til modregning, når en borger skylder det offentlige penge. Det betyder, at hvis du fx har betalt for meget i skat og derfor skal have penge tilbage, kan pengene bruges til at dække det beløb, du skylder den offentlige myndighed, fx kommunen eller Udbetaling Danmark.

Når du ikke længere skylder os penge, vil registreringen hos Gældsstyrelsen blive slettet. Hvis du vil vide, hvilke oplysninger Gældsstyrelsen har om dig, skal du kontakte Gældsstyrelsen."

(K3) "Hvis du vil klage
Du har mulighed for at klage over afgørelsen. Det gør du ved at ringe til os eller sende din klage digitalt på www.borger.dk/pension-klage. Du kan også sende din klage til Udbetaling Danmark, Pension, Kongens Vænge 8, 3400 Hillerød, gerne med titlen ''Klage over pension''. 

Vi skal have din klage, senest 4 uger efter du har modtaget afgørelsen. Så vurderer vi sagen igen. Hvis vi ikke giver dig ret i din klage, sender vi den videre til Ankestyrelsen."

(L1)"Lovgrundlag
Afgørelsen er truffet på grundlag af:"
Udbetaling Danmark-loven § 4
Pensionsloven §§ 14b, 14c,  41, 42 og § 72d, stk. 2.
(L2)"Du kan finde lovgrundlaget på www.retsinformation.dk."
"""


def create_merge_field(parent, key):
    """Create a MERGEFIELD in Word XML format"""
    # Create begin field character
    begin = OxmlElement("w:fldChar")
    begin.set(qn("w:fldCharType"), "begin")
    parent.append(begin)

    # Create instruction text
    instr = OxmlElement("w:instrText")
    instr.set(qn("xml:space"), "preserve")
    instr.text = f" MERGEFIELD {key} "
    parent.append(instr)

    # Create end field character
    end = OxmlElement("w:fldChar")
    end.set(qn("w:fldCharType"), "end")
    parent.append(end)


def create_if_field(parent, condition_key, true_result_key):
    """
    Create a nested IF field in Word XML:
    { IF "J" = "{ MERGEFIELD <condition_key> }" "{ MERGEFIELD <true_result_key> }" }
    """
    # Begin IF field
    begin = OxmlElement("w:fldChar")
    begin.set(qn("w:fldCharType"), "begin")
    parent.append(begin)

    # Instruction text for IF field
    instr = OxmlElement("w:instrText")
    instr.set(qn("xml:space"), "preserve")
    # The field code, with nested merge fields as placeholders
    instr.text = f' IF "J" = "'
    parent.append(instr)

    # Nested merge field for condition
    create_merge_field(parent, condition_key)

    # Continue IF field code
    instr2 = OxmlElement("w:instrText")
    instr2.set(qn("xml:space"), "preserve")
    instr2.text = f'" "'
    parent.append(instr2)

    # Nested merge field for true-result
    create_merge_field(parent, true_result_key)

    # End IF field code
    instr3 = OxmlElement("w:instrText")
    instr3.set(qn("xml:space"), "preserve")
    instr3.text = '"'
    parent.append(instr3)

    # End IF field
    end = OxmlElement("w:fldChar")
    end.set(qn("w:fldCharType"), "end")
    parent.append(end)


def create_merge_field_with_formatting(parent, key, run=None):
    """Create a MERGEFIELD in Word XML format, copying formatting from the original run if provided"""
    # Create begin field character
    begin = OxmlElement("w:fldChar")
    begin.set(qn("w:fldCharType"), "begin")
    parent.append(begin)

    # Create instruction text
    instr = OxmlElement("w:instrText")
    instr.set(qn("xml:space"), "preserve")
    instr.text = f" MERGEFIELD {key} "
    parent.append(instr)

    # Create end field character
    end = OxmlElement("w:fldChar")
    end.set(qn("w:fldCharType"), "end")
    parent.append(end)


def create_if_field_with_formatting(parent, condition_key, true_result_key, run=None):
    """
    Create a nested IF field in Word XML, copying formatting from the original run if provided
    """
    # Begin IF field
    begin = OxmlElement("w:fldChar")
    begin.set(qn("w:fldCharType"), "begin")
    parent.append(begin)

    # Instruction text for IF field
    instr = OxmlElement("w:instrText")
    instr.set(qn("xml:space"), "preserve")
    instr.text = f' IF "J" = "'
    parent.append(instr)

    # Nested merge field for condition
    create_merge_field_with_formatting(parent, condition_key, run)

    # Continue IF field code
    instr2 = OxmlElement("w:instrText")
    instr2.set(qn("xml:space"), "preserve")
    instr2.text = f'" "'
    parent.append(instr2)

    # Nested merge field for true-result
    create_merge_field_with_formatting(parent, true_result_key, run)

    # End IF field code
    instr3 = OxmlElement("w:instrText")
    instr3.set(qn("xml:space"), "preserve")
    instr3.text = '"'
    parent.append(instr3)

    # End IF field
    end = OxmlElement("w:fldChar")
    end.set(qn("w:fldCharType"), "end")
    parent.append(end)


def process_paragraph(paragraph, text, mappings):
    """Process a paragraph and replace text with merge fields and IF fields"""
    # Sort mappings by length of title (longest first) to prevent partial matches
    sorted_mappings = sorted(mappings.items(), key=lambda x: len(x[0]), reverse=True)

    # Clear the paragraph
    p = paragraph._p
    for child in list(p):
        p.remove(child)

    # Detect IF Betingelse lines
    if text.strip().startswith("IF Betingelse "):
        # Extract the condition title
        condition_title = text.strip()[len("IF Betingelse ") :]
        condition_key = mappings.get(condition_title)

        # Define special mappings for certain condition keys to their Html: counterparts
        special_html_mappings = {
            "ab-ubegraenset-fuldmagt": "Html:x-fuldmagtsbetingelse",
            # Add more special mappings here as needed
        }

        # Try to find true_result_key using various strategies
        true_result_key = None

        # 1. Check if there's a special mapping for this condition key
        if condition_key and condition_key in special_html_mappings:
            true_result_key = special_html_mappings[condition_key]

        # 2. Look for a key with format "Html:<condition_key>"
        elif condition_key:
            if f"Html:{condition_key}" in mappings.values():
                # Find the key by value
                for k, v in mappings.items():
                    if v == f"Html:{condition_key}":
                        true_result_key = v
                        break

        # 3. Look for a key that starts with "Html:" and contains the suffix of condition_key
        if not true_result_key and condition_key:
            suffix = None
            if "-" in condition_key:
                suffix = condition_key.split("-", 1)[1]
            elif ":" in condition_key:
                suffix = condition_key.split(":", 1)[1]

            if suffix:
                for k, v in mappings.items():
                    if v.startswith("Html:") and suffix in v:
                        true_result_key = v
                        break

        # If we still don't have a true_result_key but we have a condition_key,
        # use a default pattern as fallback
        if (
            not true_result_key
            and condition_key
            and condition_title == "Ubegrænset fuldmagt"
        ):
            true_result_key = "Html:x-fuldmagtsbetingelse"

        if condition_key and true_result_key:
            run = paragraph.add_run()
            r = run._r
            create_if_field(r, condition_key, true_result_key)
        else:
            # If mapping not found, just add the original text
            paragraph.add_run(text)
        return paragraph

    remaining_text = text
    while remaining_text:
        match_found = False

        for titel, nogle in sorted_mappings:
            idx = remaining_text.find(titel)
            if idx != -1:
                # Add text before the match
                if idx > 0:
                    paragraph.add_run(remaining_text[:idx])
                # Add the merge field
                run = paragraph.add_run()
                r = run._r
                create_merge_field(r, nogle)
                # Continue with the rest of the text
                remaining_text = remaining_text[idx + len(titel) :]
                match_found = True
                break

        if not match_found:
            # No more matches, add the remaining text
            paragraph.add_run(remaining_text)
            break

    return paragraph


def create_document_with_merge_fields(template_text, mappings):
    """Create a Word document with merge fields from template text"""
    doc = Document()

    # Split the template text into paragraphs
    paragraphs = template_text.split("\n")

    # For debugging: Keep track of replacements to show raw representation
    debug_output = []

    for para_text in paragraphs:
        if para_text.strip():
            # IF Betingelse debug preview
            if para_text.strip().startswith("IF Betingelse "):
                condition_title = para_text.strip()[len("IF Betingelse ") :]
                condition_key = mappings.get(condition_title)

                # Use the same special mappings logic as in process_paragraph
                special_html_mappings = {
                    "ab-ubegraenset-fuldmagt": "Html:x-fuldmagtsbetingelse",
                    # Add more special mappings here as needed
                }

                true_result_key = None
                if condition_key and condition_key in special_html_mappings:
                    true_result_key = special_html_mappings[condition_key]
                elif condition_key:
                    # 2. Look for a key with format "Html:<condition_key>"
                    if f"Html:{condition_key}" in mappings.values():
                        # Find the key by value
                        for k, v in mappings.items():
                            if v == f"Html:{condition_key}":
                                true_result_key = v
                                break
                    # 3. Look for a key that starts with "Html:" and contains the suffix of condition_key
                    if not true_result_key:
                        suffix = None
                        if "-" in condition_key:
                            suffix = condition_key.split("-", 1)[1]
                        elif ":" in condition_key:
                            suffix = condition_key.split(":", 1)[1]

                        if suffix:
                            for k, v in mappings.items():
                                if v.startswith("Html:") and suffix in v:
                                    true_result_key = v
                                    break

                # Special case for "Ubegrænset fuldmagt"
                if (
                    not true_result_key
                    and condition_key
                    and condition_title == "Ubegrænset fuldmagt"
                ):
                    true_result_key = "Html:x-fuldmagtsbetingelse"

                if condition_key and true_result_key:
                    debug_para = f'{{ IF "J" = "{{ MERGEFIELD {condition_key} }}" "{{ MERGEFIELD {true_result_key} }}" }}'
                else:
                    debug_para = para_text
                debug_output.append(debug_para)
            else:
                # Regular debug preview for merge fields
                debug_para = para_text
                for titel, nogle in sorted(
                    mappings.items(), key=lambda x: len(x[0]), reverse=True
                ):
                    if titel in debug_para:
                        debug_para = debug_para.replace(
                            titel, f"{{ MERGEFIELD {nogle} }}"
                        )
                debug_output.append(debug_para)

            # Add to document with actual merge fields/IF fields
            p = doc.add_paragraph()
            process_paragraph(p, para_text, mappings)

    return doc, "\n".join(debug_output)


def process_docx_template(doc, mappings):
    """Process an uploaded Word document and replace text with merge fields and IF fields while preserving formatting"""
    # Sort mappings by length of title (longest first) to prevent partial matches
    sorted_mappings = sorted(mappings.items(), key=lambda x: len(x[0]), reverse=True)

    # For debugging: Keep track of replacements to show raw representation
    debug_output = []

    for para in doc.paragraphs:
        para_text = para.text
        if not para_text.strip():
            continue

        # Debug output handling
        if para_text.strip().startswith("IF Betingelse "):
            # Handle debug output for IF condition
            condition_title = para_text.strip()[len("IF Betingelse ") :]
            condition_key = mappings.get(condition_title)
            special_html_mappings = {
                "ab-ubegraenset-fuldmagt": "Html:x-fuldmagtsbetingelse"
            }
            true_result_key = None

            # Find the true_result_key using existing logic
            if condition_key and condition_key in special_html_mappings:
                true_result_key = special_html_mappings[condition_key]
            elif condition_key:
                if f"Html:{condition_key}" in mappings.values():
                    for k, v in mappings.items():
                        if v == f"Html:{condition_key}":
                            true_result_key = v
                            break
                if not true_result_key:
                    suffix = None
                    if "-" in condition_key:
                        suffix = condition_key.split("-", 1)[1]
                    elif ":" in condition_key:
                        suffix = condition_key.split(":", 1)[1]
                    if suffix:
                        for k, v in mappings.items():
                            if v.startswith("Html:") and suffix in v:
                                true_result_key = v
                                break
            if (
                not true_result_key
                and condition_key
                and condition_title == "Ubegrænset fuldmagt"
            ):
                true_result_key = "Html:x-fuldmagtsbetingelse"

            if condition_key and true_result_key:
                debug_para = f'{{ IF "J" = "{{ MERGEFIELD {condition_key} }}" "{{ MERGEFIELD {true_result_key} }}" }}'
            else:
                debug_para = para_text
            debug_output.append(debug_para)

            # Special handling for IF Betingelse paragraphs
            # Keep the first run's formatting (or create a new one if none exist)
            if len(para.runs) > 0:
                # Clear the paragraph but keep its formatting properties
                original_run = para.runs[0]
                for run in list(para.runs):
                    run._element.getparent().remove(run._element)

                # Create a new run with the same formatting
                new_run = para.add_run()
                # Copy formatting from original run
                new_run.bold = original_run.bold
                new_run.italic = original_run.italic
                new_run.underline = original_run.underline
                if original_run.font.name:
                    new_run.font.name = original_run.font.name
                if original_run.font.size:
                    new_run.font.size = original_run.font.size
                if original_run.font.color.rgb:
                    new_run.font.color.rgb = original_run.font.color.rgb

                # Add the IF field with preserved formatting
                if condition_key and true_result_key:
                    create_if_field_with_formatting(
                        new_run._element, condition_key, true_result_key, original_run
                    )
                else:
                    new_run.text = para_text
            else:
                # No runs, create a new one
                run = para.add_run()
                if condition_key and true_result_key:
                    create_if_field_with_formatting(
                        run._element, condition_key, true_result_key
                    )
                else:
                    run.text = para_text
        else:
            # Regular text paragraph - handle at run level
            # First collect all runs and their text
            run_data = [(run, run.text) for run in para.runs]

            # Create debug preview
            debug_para = para_text
            for titel, nogle in sorted_mappings:
                if titel in debug_para:
                    debug_para = debug_para.replace(titel, f"{{ MERGEFIELD {nogle} }}")
            debug_output.append(debug_para)

            # Process each run, replacing text with merge fields where needed
            for i, (run, run_text) in enumerate(run_data):
                if not run_text.strip():
                    continue

                # Check for replacements in this run
                replaced = False
                for titel, nogle in sorted_mappings:
                    if titel in run_text:
                        # This run contains text that needs to be replaced
                        # Keep the original run for its formatting
                        original_format = run._element

                        # Replace the run's text with parts before, after, and merge field
                        parts = run_text.split(titel)

                        # Clear this run's content
                        run._element.clear_content()

                        # Add text before the field
                        if parts[0]:
                            run.text = parts[0]

                        # Add the merge field
                        create_merge_field_with_formatting(run._element, nogle, run)

                        # If there's text after the field, create a new run with the same formatting
                        if len(parts) > 1 and parts[1]:
                            new_run = para.add_run(parts[1])
                            # Copy formatting
                            new_run.bold = run.bold
                            new_run.italic = run.italic
                            new_run.underline = run.underline
                            if run.font.name:
                                new_run.font.name = run.font.name
                            if run.font.size:
                                new_run.font.size = run.font.size
                            if run.font.color.rgb:
                                new_run.font.color.rgb = run.font.color.rgb

                        replaced = True
                        break

                # If no replacements were made, keep the run as is
                if not replaced:
                    pass  # No changes needed

    return doc, "\n".join(debug_output)


# Try to load default mapping at startup
DEFAULT_MAPPING_PATH = os.path.join("documents", "liste over alle nøgler.xlsx")
default_mappings = None
if os.path.exists(DEFAULT_MAPPING_PATH):
    try:
        df_default = pd.read_excel(DEFAULT_MAPPING_PATH, sheet_name="query")
        if "Titel" in df_default.columns and "Nøgle" in df_default.columns:
            default_mappings = {
                row["Titel"]: row["Nøgle"] for _, row in df_default.iterrows()
            }
    except Exception as e:
        default_mappings = None

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
    try:
        df = pd.read_excel(uploaded_file, sheet_name="query")
        if "Titel" not in df.columns or "Nøgle" not in df.columns:
            st.error(
                "Excel-filen skal indeholde kolonnerne 'Titel' og 'Nøgle' i arket 'query'."
            )
        else:
            st.write("Antal koblinger fundet:", df.shape[0])
            with st.expander("Vis titel/nøgle-par (fra uploadet fil)", expanded=False):
                st.dataframe(df[["Titel", "Nøgle"]], hide_index=True)
            mappings = {row["Titel"]: row["Nøgle"] for _, row in df.iterrows()}
    except Exception as e:
        st.error(f"Fejl ved behandling af Excel-fil: {str(e)}")
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
    
    2. **Upload en Word-skabelon**. Skabelonen skal indeholde de tekststrenge, der står i "Titel"-kolonnen
    
    3. **Generér dokumentet** ved at klikke på knappen
    
    4. **Download** det færdige Word-dokument med fletfelter
    
    Appen udskifter hver forekomst af tekst fra "Titel"-kolonnen med tilsvarende Word-fletfelter ud fra "Nøgle"-værdierne.
    """
    )
