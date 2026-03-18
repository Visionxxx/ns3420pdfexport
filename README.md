# ns3420pdfexport

Konverterer NS 3420 byggebeskrivelse-PDF til regneark (Excel/CSV), NS 3459 XML og JSON.


## Installasjon

```bash
pip install pymupdf openpyxl
```

## Bruk

```bash
# Standard: eksporter alle formater (xlsx, csv, xml, json)
python ns3420pdfimport.py beskrivelse.pdf

# Kun Excel
python ns3420pdfimport.py beskrivelse.pdf --format xlsx

# CSV med egendefinert filnavn
python ns3420pdfimport.py beskrivelse.pdf --format csv --output resultat

# Detaljert output
python ns3420pdfimport.py beskrivelse.pdf --verbose

# Fra pre-ekstrahert tekstfil (fallback)
python ns3420pdfimport.py ekstrahert.txt --format xlsx
```

## Eksportformater

| Format | Beskrivelse |
|--------|-------------|
| **xlsx** | Excel med 5 ark: Postliste, Prismatrise, Kapitteloversikt, Full beskrivelse, Statistikk |
| **csv** | Semikolon-separert, ns3420reader-kompatibel import |
| **xml** | NS 3459-format |
| **json** | Strukturert JSON for integrasjon |

### Kolonner (ns3420reader-format)

```
Postnr ; NS3420 ; Emne ; Beskrivelse ; Mengde ; Enhet ; Pris ; Sum
```

- **Postnr**: Hierarkisk postnummer (f.eks. `05.21.2`, `09.235.10.1.1`)
- **NS3420**: NS 3420 kode (f.eks. `LB1.1112A`, `WZA`)
- **Emne**: Seksjonstittel
- **Beskrivelse**: Full spesifikasjonstekst inkl. alle detaljer, lokalisering, andre krav
- **Mengde**: Kvantum
- **Enhet**: m, m2, m3, stk, kg, tonn, RS
- **Pris / Sum**: Tomme (fylles av tilbyder)

## Teknisk

### Koordinat-basert PDF-parsing (standard)

Bruker PyMuPDF til å lese PDF-en med nøyaktige x/y-koordinater for hver tekstspan. Kolonnetilordning basert på faste x-posisjoner som er konsistente i NS 3420-dokumenter generert av ISY Beskrivelse:

| Kolonne | x-område |
|---------|----------|
| Postnr | 39 - 99 |
| NS-kode/Spesifikasjon | 99 - 365 |
| Enh. | 365 - 397 |
| Mengde | 397 - 460 |
| Pris | 460 - 524 |
| Sum | 524 - 580 |

Superscript-enheter (m², m³) detekteres via fontstørrelse (7pt vs 10pt).

Split postnumre (lange numre som brytes over tabellrader) sammenslås automatisk.

### Tekst-basert parsing (fallback)

For `.txt`-filer brukes regex-basert parsing av ekstrahert tekst.

## Avhengigheter

- **pymupdf** (fitz) - PDF-parsing med koordinater (anbefalt)
- **openpyxl** - Excel-eksport

Valgfritt for fallback tekst-ekstraksjon:
- **pdfplumber** - Alternativ PDF-tekstekstraksjon
