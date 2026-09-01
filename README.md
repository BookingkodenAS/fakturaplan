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
  `skills/fakturaoppdatering/SKILL.md` for full rutine og spesialtilfeller.
- `faktura_kalkulator.cpython310.pyc` – **kompilert bytecode, ikke kildekode.**
  Beregner fakturabeløp per kunde ut fra prismodellene i Fakturaplan. Kilden
  (`faktura_kalkulator.py`) lå ikke i det som ble delt fra Cowork-prosjektet,
  kun denne kompilerte `.pyc`-filen (Python 3.10). Automatisk dekompilering ble
  forsøkt, men ga ufullstendig og delvis **feil** logikk (bl.a. en gal
  grensesjekk i oppslagsfunksjonen), så den er ikke lagt inn her – det ville
  vært risikabelt å basere fakturaberegning på gjettet kode. **Kildekoden bør
  legges inn i `scripts/faktura_kalkulator.py` så snart den er tilgjengelig**,
  f.eks. fra `C:\Users\hei\Claude\Projects\Fakturaplan\faktura_kalkulator.py`
  på maskinen der Cowork kjører.

### `skills/fakturaoppdatering/`
- `SKILL.md` – rutinebeskrivelsen som styrer den månedlige
  fakturaoppdateringen i Cowork (kjøres normalt automatisk den 1. i hver
  måned).

## Status / TODO
- [ ] Legg til kildekoden for `faktura_kalkulator.py` (se over)
- [ ] Vurder om `.pyc`-filen skal fjernes når kildekoden er på plass
