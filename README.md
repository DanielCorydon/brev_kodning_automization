# Brevkoder Automatisering

Dette projekt er en Streamlit-app, der automatiserer indsættelse af fletfelter i Word-dokumenter baseret på koblinger fra en Excel-fil. Appen er designet til at hjælpe med at generere breve, hvor bestemte tekststrenge automatisk udskiftes med Word-fletfelter, hvilket gør det nemt at lave masseudsendelser eller tilpasse dokumenter.

## Funktioner
- **Upload af Excel-fil**: Indlæs en fil med koblinger mellem tekststrenge ("Titel") og fletfeltnavne ("Nøgle").
- **Upload af Word-skabelon**: Indlæs en Word-skabelon, hvor tekststrenge fra Excel-filen erstattes med fletfelter.
- **Automatisk generering**: Dokumentet genereres automatisk, når begge filer er uploadet.
- **Download**: Download det færdige Word-dokument med indsatte fletfelter.
- **Debug/Preview**: Se en forhåndsvisning af, hvordan fletfelterne indsættes.

## Sådan bruges appen
1. **Upload en Excel-fil** med et ark kaldet `query` og to kolonner:
    - `Titel`: Tekststrenge, der optræder i dit brev
    - `Nøgle`: Felt-navne, der skal bruges som Word-fletfelter
2. **Upload en Word-skabelon**. Skabelonen skal indeholde de tekststrenge, der står i "Titel"-kolonnen.
3. **Download** det færdige Word-dokument med fletfelter.

Appen udskifter hver forekomst af tekst fra "Titel"-kolonnen med tilsvarende Word-fletfelter ud fra "Nøgle"-værdierne.

## Krav
- Python 3.8+
- [Streamlit](https://streamlit.io/)
- [pandas](https://pandas.pydata.org/)
- [python-docx](https://python-docx.readthedocs.io/)

Alle afhængigheder kan installeres via `pdm` eller `pip`:

```sh
pdm install
# eller
pip install -r requirements.txt
```

## Kørsel
Kør appen med:

```sh
streamlit run app.py
```

## Projektstruktur
```
app.py                  # Hovedapplikation (Streamlit)
documents/              # Mapper til eksempelfiler og skabeloner
src/components/         # (Evt. komponenter til udvidelse)
pyproject.toml, pdm.lock# Afhængigheder og projektopsætning
```

## Licens
Dette projekt er kun til privat brug og uddannelsesformål.
