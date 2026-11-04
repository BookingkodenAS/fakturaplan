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
- Willis Towers Watson AS (linje 77): fast kr 490/mnd tillegg for bestilt samtaleoverføringsrapport. Gjaldt IKKE for august (bekreftet av Geir 2026-09-17). **Lagt til i I92-formelen fra og med september 2026-kjøringen (2026-10-01)** — `+490` er nå en permanent del av formelen, ikke fjern den.
- Enova Trondheim (linje 65): kvartalsfakturering (jan/apr/jul/okt, forskudd). August 2026 var bekreftet IKKE faktureringsmåned (I76 tømt den perioden). **September 2026-kjøringen (2026-10-01) satte H76 til Q3-kumulativt jul+aug+sep = 28+38+33 = 99, og gjenopprettet formelen `=5140+MAX(0,H76-200)*36`** (= 5140 kr, ingen overpris siden 99<200). Neste steg: ved oktoberkjøringen (prosesserer oktober-data, ~1. nov) skal H76 nullstilles til KUN oktober sin verdi (start på nytt Q4-kvartal: okt+nov+des), og I76 tømmes igjen til neste forskuddsfaktura (oktober var allerede fakturamåned, så Q4 forskuddet er betalt — neste faktureringspunkt er januar 2027 med Q4-overpris-oppgjør).
- NYD – Norsk Yrkesdykkerskole (linje 67): fakturerer forskudd pr. kvartal (feb/mai/aug/nov), flat kr 10 230 uten trapp. **2026-11-04: lagt om fra engangsordre-per-kvartal til én løpende abonnementsordre i Tripletex** (ordre 28406, kundenr 30308), så det trengs ikke lenger en ny manuell ordre hvert kvartal.
- Kinnarps (linje 79) og T. Solbakken Trafikkskole (linje 80): samme modell som Auto Vest (kr 1500 fast + kr 950 per påbegynt 20-steg). **2026-11-04: status satt til Aktiv, begge satt opp med løpende abonnementsordre i Tripletex** (Kinnarps: ordre 28404, kundenr 33203; T. Solbakken: ordre 28405, kundenr 33201). Begge fantes fra før som Tripletex-kunder med én lukket etableringsordre fra august 2026 – den er IKKE rørt.
- Auto Vest (linje 82): kr 1500 fast + kr 950/20-steg. Lagt inn 2026-11-04 (org.nr 932 540 568, Varden 25, 5535 Haugesund), løpende abonnement i Tripletex (ordre 28403, kundenr 33204).
- Fit4 AS (linje 70) + Fit4 e-post (linje 81): **2026-11-04 bekreftet ny modell av Geir**: kr 3292/mnd inkl. 100 ENHETER TOTALT (telefon H85 + e-post H103 samlet), + kr 28/enhet over 100, hele summen ×1,1. H103 (e-post) oppdateres manuelt hver måned på linje 81. Tripletex-ordren (27278) er lagt om til denne modellen – den gamle separate flate e-postlinjen (kr 1548) er nullstilt, IKKE slett den linjen på nytt (sletting av denne linjen feilet stille to ganger i Tripletex, satte price=0 i stedet).
- IKT Nor AS (linje 39): grunnprisformelen (kr 860 + kr 42/anrop over 20) er den RIKTIGE kilden til sannhet, bekreftet av Geir 2026-11-04 – Tripletex-ordren hadde drevet seg bort fra formelen (feil fastlinje), rettet til å matche formelen igjen.
- Klinikk Øre-nese-hals AS (linje 56): **2026-11-04: ny prismodell** fra we4you-tilbudet "Tilbud på kundesentertjeneste" (fast beredskap kr 1500 inntil 30 anrop, deretter 20-trinns trapp til kr 37930 ved 891–910 anrop, ×1,1) – erstatter den gamle CEG-tabell-formelen som ga feil tall. Tabellen ligger i eget ark `ONH tabell` i Prisoversikt-filen. Tripletex-ordren (27099) er oppdatert til ett linjebeløp fra denne tabellen.
- Element Logic AS (linje 9): kvartalsfaktura forskuddsvis (jan/apr/jul/okt), kr 8880/kvartal ×1,1 = kr 9768. Tripletex-ordren (27410) var allerede riktig satt opp som `subscriptionPeriodsOnInvoice=3, ADVANCE`, men linjeprisen var feilaktig satt til kun 1/3 av kvartalsbeløpet (2960) – rettet 2026-11-04 til 9768. Element Logic Sweden AB kjøres fortsatt som egen engangsordre per kvartal (mar/jun/sep/des) à kr 8880 (uten 10%-økning ennå, ikke rørt) – IKKE bland sammen med Element Logic AS sin ordre.
- Veidekke Industri AS – Møre (linje 27, kundenr 32722): **ordre 28326 er den gjeldende** i Tripletex (bekreftet av Geir 2026-11-04) – IKKE ordre 27123, som også finnes åpen på samme kunde men er en duplikat/gammel ordre.
- Veidekke Industri AS – Sogn (linje 28, kundenr 33197, oppstart mars 2026): egen Tripletex-kunde (ikke samme som Møre), ordre 27827.
- Spraying Systems Norge AS (linje 22): **rettet fra feilaktig "Spraying Systems Sverige AB"** i Prisoversikten 2026-11-04. Ordre 27116 er den gjeldende i Tripletex (ordre 28325 er en gammel duplikat, ikke oppdatert).
- Enova Trondheim (Compass, linje 65) og Norsk Medisinsk Syklotronsenter (linje 66/77): **2026-11-04: periodene rullet videre** til okt–des 2026 på begge Tripletex-ordrene (27086: kr 5654 forskuddsvis; 27107: kr 3157 etterskuddsvis). Neste rulling: 27086 til jan–mar 2027 ved januar-kjøringen (forskudd), 27107 til jan–mar 2027 ved mars-kjøringen (etterskudd).

### Tripletex-synkronisering (etablert 2026-11-04)

Hver åpne sentralbordkunde-ordre i Tripletex har department "Svartjeneste"
(id 180801), manuell fakturering (`isSubscriptionAutoInvoicing: false`), og
som regel 1-2 linjer: en fast/strukturell linje + en "justeringslinje" som
brukes til å treffe månedens Prisoversikt-beløp nøyaktig. Konvensjon for
justeringslinjen: beskrivelse som "Klikk pris", "Besvarte anrop" e.l., og
**pris nøyaktig 1 eller 0** signaliserer at det er en ren plugg-linje (kan
settes til hvilket som helst tall, også negativt, for å treffe riktig
total) – IKKE en reell per-enhets-rate. En linje med en ekte per-enhets-pris
(f.eks. 42 kr/anrop) og en reell qty (antall anrop over grensen) skal
justeres via QUANTITY, ikke pris. Tripletex tillater ikke `quantity: 0` på
en prissatt linje – sett i stedet `quantity: 1, unitPriceExclVat: 0` for å
nullstille en slik linje midlertidig. `deleteLineIds` har i denne økten vist
seg å feile stille (rapporterer "deleted" men linjen består) – bruk price=0
som trygg erstatning for sletting.

Flere kunder har to åpne ordre på samme kundenummer (historiske duplikater)
– sjekk alltid med Geir hvilken som er "gjeldende" før du oppdaterer, flagg
heller enn å gjette.

## Steg 4 – Ukjente køer

Hvis scriptet rapporterer ukjente køer (ny kunde ikke i mappingen i
`oppdater_h_kolonne.py` eller `faktura_kalkulator.py` sin `KUNDER`-liste),
IKKE gjett – flagg dette tydelig til Geir slik at han kan legge til riktig
linjenummer/kundekonfigurasjon.

## Steg 5 – Skriv endringene (etter bekreftelse fra Geir)

```bash
python3 scripts/oppdater_h_kolonne.py "data/zisson/<zisson-fil>.xlsx" --skriv
```

**Ikke bruk `faktura_kalkulator.py` til å beregne fakturagrunnlaget** – det er
en ufullstendig separat Python-reimplementasjon (mangler bl.a. Gass og
Pusteservice AS, teller Norenco feil, ingen NEMUS-fordeling per klinikk). Fra
og med august 2026 har `PrisoversiktSentralbord2026.xlsx` egne formler
(I-kolonnen, VLOOKUP mot `Tabeller` og de andre referansearkene) som er kilde
til sannhet. openpyxl beregner ikke formler, så etter `--skriv` må filen
kjøres gjennom LibreOffice for å få friske beregnede verdier i I-kolonnen:

```bash
soffice --headless --norestore -env:UserInstallation=file:///tmp/lo_$$ \
  --convert-to xlsx --outdir /tmp/recalc data/PrisoversiktSentralbord2026.xlsx
cp /tmp/recalc/PrisoversiktSentralbord2026.xlsx data/PrisoversiktSentralbord2026.xlsx
```

(Hvis `soffice`/`libreoffice-calc` ikke er installert: `apt-get update && apt-get install -y libreoffice-calc`.)

Les deretter summen av I4:I93 (samme range som `BEREGNET TOTAL`-raden bruker)
og list opp linjer med tom I-verdi (kvartalsvise/manuelle poster som ikke er
fakturamåned, eller nye kunder uten prismodell ennå) – disse skal IKKE telles
med i totalen.

```bash
git add data/PrisoversiktSentralbord2026.xlsx
git commit -m "Fakturaoppdatering <måned> <år>"
git push origin main
```

## Steg 6 – Rapporter til Geir

Avslutt alltid med en kort rapport på norsk: status på Zisson-rapporten,
hvilke H-kolonne-endringer som ble foreslått/skrevet, avvik i
spesialtilfellene, ukjente køer som må avklares, og NEMUS-påminnelsen.
Skriv aldri `--skriv` uten at Geir har bekreftet tørrkjøringen først.
