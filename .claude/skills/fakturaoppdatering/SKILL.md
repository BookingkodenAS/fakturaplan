---
name: fakturaoppdatering
description: >
  Oppdaterer H-kolonnen (Anrop besvart) i Prisoversikt-Sentralbord for ny
  faktureringsperiode basert paa Zisson maanesrapport fra geir@we4you.no sin
  innboks, og beregner fakturagrunnlag per kunde. Bruk ALLTID denne skillen
  naar Geir laster opp en ny Zisson maanedsstatistikk og ber om aa oppdatere
  anropstall, starte ny fakturering, oppdatere H-kolonnen, kjore
  anropsoppdatering, eller klargjore prisoversikten for neste maaned.
  Trigges ogsaa naar Geir sier "ny maanesrapport", "oppdater anrop",
  "fakturering neste maaned", eller ligner, og av den maanedlige Routine som
  kjorer 1. i hver maaned.
---

# Fakturaoppdatering – Anrop fra Zisson til Prisoversikt og fakturagrunnlag

Dette er den flyttede Cowork-rutinen for We4you AS sitt sentralbord, tilpasset
til å kjøre i dette Git-repoet (`fakturaplan`) i stedet for den lokale
Cowork-mappen `C:\Users\hei\Claude\Projects\Fakturaplan\`.

## Repostruktur

```
data/
  PrisoversiktSentralbord2026.xlsx   – prisoversikt, H-kolonnen oppdateres her
  Fakturaplan2026.xlsx               – termintabell/kvartalsfakturering
  Kvartalsvis_Terminfakturering_Sentralbordet.xlsx
  NEMUS_Anrop_Fordeling_Juni2026.xlsx (og senere måneders NEMUS-fordeling)
scripts/
  oppdater_h_kolonne.py   – oppdaterer H-kolonnen i Prisoversikten
  faktura_kalkulator.py   – beregner fakturabeløp per kunde (Fakturagrunnlag_<måned>.xlsx)
```

## Steg 0 – Finn Zisson-rapporten

Zisson sender en e-post fra `noreply@zisson.com` med emne
`Zisson statistics We4you- Måned (4830) ÅÅÅÅ-MM-DD.xlsx` til
`geir@we4you.no` (og kjetil@we4you.no, kundeservice@we4you.no) rett etter
midnatt den 1. i måneden. Rapporten havner også automatisk i SharePoint under
`We4you/Finans/Fakturering/Fakturering Sentralbord/Fakturering 2026/Statistikk/`.

1. Søk i `geir@we4you.no` sin innboks (`outlook_email_search` med
   `mailboxOwnerEmail: geir@we4you.no`, `sender: zisson.com`,
   `order: newest`) for å bekrefte at rapporten for perioden har kommet inn.
2. **Kjent begrensning:** Microsoft 365-connectoren i denne økten kan bekrefte
   at e-posten og vedlegget finnes (`read_resource` på meldingen), men kan
   **ikke laste ned selve det binære vedlegget** ("Binary attachment — content
   cannot be returned inline"). Å lese samme rapport fra SharePoint via
   `sharepoint_search` + `read_resource` gir kun en **delvis** tekstdump av
   store ark (Kødetaljer-arket har ofte 60-70+ køer/1300+ rader, og
   uttrekket kutter etter noen hundre rader) — IKKE nok til å stole på for en
   reell fakturakjøring.
3. Derfor: be Geir om å laste opp/lime inn selve `.xlsx`-filen i chatten
   (samme som han allerede har gjort tidligere), eller lagre den til
   prosjektmappen manuelt hvis Cowork-mappen fortsatt brukes parallelt.
   Fortsett IKKE til steg 2 med ufullstendige data.

## Steg 1 – Legg Zisson-filen i repoet

Når filen er mottatt (opplastet i chat), kopier den inn i en midlertidig
katalog i klonen, f.eks. `data/zisson/We4you- Måned (XXXX) ÅÅÅÅ-MM-DD.xlsx`
(denne mappen commit'es ikke nødvendigvis inn i git — det holder at scriptet
finner den lokalt for kjøringen).

## Steg 2 – Tørrkjøring (IKKE skriv ennå)

```bash
python3 scripts/oppdater_h_kolonne.py "data/zisson/<zisson-fil>.xlsx"
```

Les output grundig og noter alle foreslåtte endringer i H-kolonnen (Anrop
besvart) i `data/PrisoversiktSentralbord2026.xlsx`.

## Steg 3 – Kontroller spesialtilfellene mot outputen

- Dal Media (linje 8): sum av 4 køer – Dal Media + TDC Caravan + Saltando + Victoria Vottestad
- Norenco (linje 17): inkluderer Mekan-køen
- Standard Norge (linje 20): inkluderer Standard Online
- KTV Group (linje 15): inkluderer KTV Kveld og helg
- Safenordic (linje 64): inkluderer Touchcom
- Gulvfag (linje 32): inkluderer Nye Gulv Norge + Adler parkettsliperi
- Fit4 (linje 70): inkluderer Helselaben
- Veidekke e-post (linje 28): H-kolonnen er antall e-poster, oppdateres manuelt – IKKE fra Zisson
- Enova Trondheim (linje 65): kvartalsfakturering – sjekk om det er faktureringsmåned
- NYD (linje 67) og Feste Trafikkskole (linje 68): sjekk faktureringsdato
- NEMUS-klinikkene (linje 50–54): fordeles manuelt via NEMUS_Anrop_Fordeling-filen, scriptet oppdaterer kun totalsummen i B5

## Steg 4 – Ukjente køer

Hvis scriptet rapporterer ukjente køer (ny kunde ikke i mappingen i
`oppdater_h_kolonne.py` eller `faktura_kalkulator.py` sin `KUNDER`-liste),
IKKE gjett – flagg dette tydelig til Geir slik at han kan legge til riktig
linjenummer/kundekonfigurasjon.

## Steg 5 – Skriv endringene (etter bekreftelse fra Geir)

```bash
python3 scripts/oppdater_h_kolonne.py "data/zisson/<zisson-fil>.xlsx" --skriv
python3 scripts/faktura_kalkulator.py "data/zisson/<zisson-fil>.xlsx"
```

`faktura_kalkulator.py` bygger `Fakturagrunnlag_<måned>.xlsx`. Legg denne i
`data/fakturagrunnlag/` og commit den sammen med den oppdaterte
`PrisoversiktSentralbord2026.xlsx`, med commit-melding som nevner
faktureringsperioden.

```bash
git add data/PrisoversiktSentralbord2026.xlsx data/fakturagrunnlag/Fakturagrunnlag_<måned>.xlsx
git commit -m "Fakturaoppdatering <måned> <år>"
git push origin main
```

## Steg 6 – Rapporter til Geir

Avslutt alltid med en kort rapport på norsk: status på Zisson-rapporten,
hvilke H-kolonne-endringer som ble foreslått/skrevet, avvik i
spesialtilfellene, ukjente køer som må avklares, og NEMUS-påminnelsen.
Skriv aldri `--skriv` uten at Geir har bekreftet tørrkjøringen først.
