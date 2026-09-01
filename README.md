# Fakturaplan

Fakturaplan er We4you AS sin rutine for å gjøre den månedlige faktureringen av
sentralbordkunder enklere og delvis automatisert. Prosjektet er flyttet hit fra
Claude Cowork (lokal prosjektmappe `C:\Users\hei\Claude\Projects\Fakturaplan\`).

## Innhold

### `data/`
Grunnlagsdata og prisdokumenter:
- `Fakturaplan2026.xlsx` – hovedoversikt for faktureringsåret 2026
- `PrisoversiktSentralbord2026.xlsx` – prisoversikt per kunde, inkl. H-kolonnen
  (Anrop besvart) som oppdateres månedlig fra Zisson-statistikken
- `Kvartalsvis_Terminfakturering_Sentralbordet.xlsx` – kunder som faktureres
  kvartalsvis (f.eks. Enova Trondheim)
- `NEMUS_Anrop_Fordeling_Juni2026.xlsx` – manuell fordeling av NEMUS-anrop per
  klinikk
- `pristabell_sentralbordet_ceg_tABELL.docx`, `Pristabell_we4you_2026_kontraktsklar.docx`
  – kontraktsklare pristabeller

### `scripts/`
- `oppdater_h_kolonne.py` – leser nyeste Zisson Interact månedsstatistikk og
  oppdaterer H-kolonnen (Anrop besvart) i Prisoversikt-Sentralbord. Kjøres som
  tørrkjøring først, deretter med `--skriv` etter bekreftelse. Se
  `.claude/skills/fakturaoppdatering/SKILL.md` for full rutine og spesialtilfeller.
- `faktura_kalkulator.py` – beregner fakturabeløp per kunde ut fra
  prismodellene i Fakturaplan, basert på Zisson-statistikken. Inneholder
  kundekonfigurasjon (`KUNDER`) og prismodeller (`std`, `ceg`, egne tabeller,
  fast pris/stykkpris-varianter osv.) og bygger en formatert
  `Fakturagrunnlag_<måned>.xlsx`. Merk: standardstien i `main()` peker til en
  Cowork-arbeidsmappe (`/sessions/.../Fakturaplan/`) – oppgi Zisson-filen som
  argument når scriptet kjøres et annet sted.

### `.claude/skills/fakturaoppdatering/`
- `SKILL.md` – rutinebeskrivelsen som styrer den månedlige fakturaoppdateringen.
  Kjøres automatisk 1. i hver måned via en Claude-scheduled Routine, og kan
  også trigges manuelt ved å be Claude om "fakturaoppdatering" i dette repoet.
  **Kjent begrensning:** Zisson-rapporten kan ikke lastes ned automatisk som
  binært vedlegg fra Outlook via Microsoft 365-connectoren, og
  SharePoint-uttrekket av samme rapport blir for stort til å hentes komplett.
  Rutinen finner og bekrefter at rapporten har kommet inn, men trenger at
  Geir laster opp selve `.xlsx`-filen i chatten før tallene kan beregnes.

