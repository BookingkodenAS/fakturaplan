# Fakturaplan

Fakturaplan er We4you AS sin rutine for å gjøre den månedlige faktureringen av
sentralbordkunder enklere og delvis automatisert. Prosjektet er flyttet hit fra
Claude Cowork (lokal prosjektmappe `C:\Users\hei\Claude\Projects\Fakturaplan\`).

## Innhold

### `data/`
Grunnlagsdata og prisdokumenter:
- `Fakturaplan2026.xlsx` – hovedoversikt for faktureringsåret 2026
- `PrisoversiktSentralbord2026.xlsx` – prisoversikt per kunde. **Dette er nå
  kilden til sannhet for fakturaberegningen**: H-kolonnen (Besvarte anrop)
  oppdateres månedlig fra Zisson-statistikken, og I-kolonnen (Beregnet pris)
  beregnes av arkets egne formler (VLOOKUP mot `Tabeller`-arket og de andre
  referansearkene: `Standard 0-99`, `CEG-tabell`, `Caverion`,
  `Fysikalsk institutt`, `Promon tabell`, `ENRX tabell`, `Rambøll tabell`,
  `Kvartalsfakturering`). Siden openpyxl (som scriptene bruker) ikke
  beregner formler selv, må filen kjøres gjennom LibreOffice headless for å
  få oppdaterte tall etter en H-kolonne-endring:
  `soffice --headless --convert-to xlsx --outdir <mappe> PrisoversiktSentralbord2026.xlsx`
  (Excel gjør dette automatisk når filen åpnes, så det trengs kun for
  rapportering/verifisering i denne økten.)
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
- `faktura_kalkulator.py` – **utdatert/ufullstendig**, ikke lenger kilde til
  sannhet. Dette var en separat Python-reimplementasjon av prismodellene, men
  den mangler flere kunder (bl.a. Gass og Pusteservice AS), regner Norenco
  feil (teller ikke med Mekan-aliaset), og har ikke NEMUS-fordeling per
  klinikk. Etter at Prisoversikt-arket ble utvidet med egne
  formler/referanseark (se over) er det arkets egne formler som gjelder.
  Behold scriptet for historikk, men ikke bruk det til å generere
  fakturagrunnlag uten å kryssjekke mot Prisoversikten.

### `.claude/skills/fakturaoppdatering/`
- `SKILL.md` – rutinebeskrivelsen som styrer den månedlige fakturaoppdateringen.
  Kjøres automatisk 1. i hver måned via en Claude-scheduled Routine, og kan
  også trigges manuelt ved å be Claude om "fakturaoppdatering" i dette repoet.
  **Kjent begrensning:** Zisson-rapporten kan ikke lastes ned automatisk som
  binært vedlegg fra Outlook via Microsoft 365-connectoren, og
  SharePoint-uttrekket av samme rapport blir for stort til å hentes komplett.
  Rutinen finner og bekrefter at rapporten har kommet inn, men trenger at
  Geir laster opp selve `.xlsx`-filen i chatten før tallene kan beregnes.

