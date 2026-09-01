---
name: fakturaoppdatering
description: >
  Oppdaterer H-kolonnen (Anrop besvart) i Prisoversikt-Sentralbord for ny
  faktureringsperiode basert paa Zisson maanesrapport lastet opp i prosjektmappen.
  Bruk ALLTID denne skillen naar Geir laster opp en ny Zisson maanedsstatistikk
  og ber om aa oppdatere anropstall, starte ny fakturering, oppdatere H-kolonnen,
  kjore anropsoppdatering, eller klargjore prisoversikten for neste maaned.
  Trigges ogsaa naar Geir sier "ny maanesrapport", "oppdater anrop", "fakturering
  neste maaned", eller ligner.
---

# Fakturaoppdatering – Anrop fra Zisson til Prisoversikt

Dette skriptet leser den nyeste Zisson Interact maanesstatistikk-filen i
prosjektmappen og oppdaterer H-kolonnen (Anrop besvart) for alle kunder i
Prisoversikt-Sentralbord-*.xlsx.

## Prosjektmappe

```
C:\Users\hei\Claude\Projects\Fakturaplan\
```

## Slik kjorer du oppdateringen

### Steg 1 – Torkjoring (vis hva som vil endres)

```bash
python3 "C:\Users\hei\Claude\Projects\Fakturaplan\oppdater_h_kolonne.py"
```

Les outputen grundig. Sjekk at alle endringer ser riktige ut for den aktuelle
maneden. Spesielt:
- **Dal Media (L8)**: Summen av 4 koer (Dal Media + TDC Caravan + Saltando + Victoria Vottestad)
- **Norenco (L17)**: Inkluderer Mekan-koen
- **Standard Norge (L20)**: Inkluderer Standard Online
- **KTV Group (L15)**: Inkluderer KTV Kveld og helg
- **Safenordic (L64)**: Inkluderer Touchcom
- **Gulvfag (L32)**: Inkluderer Nye Gulv Norge + Adler parkettsliperi
- **Fit4 (L70)**: Inkluderer Helselaben

### Steg 2 – Lukk Prisoversikten i Excel

Be Geir om aa lukke/lagre Prisoversikt-Sentralbord-*.xlsx i Excel foer du
skriver endringene. (Filen maa vaere lukket for at oppdateringen skal fungere.)

### Steg 3 – Skriv endringene

```bash
python3 "C:\Users\hei\Claude\Projects\Fakturaplan\oppdater_h_kolonne.py" --skriv
```

### Steg 4 – NEMUS (linje 50–54)

NEMUS-klinikkene faktureres separat og maa fordeles manuelt:

1. Scriptet oppdaterer automatisk B5 i `NEMUS_Anrop_Fordeling_*.xlsx` med
   totalt antall anrop for Nemus (ko 4768).
2. Geir fyller inn fordeling per klinikk (Horten, Larvik, Notteroy, Sandefjord,
   Tonsberg) manuelt basert paa Teams-registrering.
3. Etter at NEMUS-filen er utfylt, oppdater H50–H54 i Prisoversikten manuelt.

### Steg 5 – Aapne filen og kontroller

Be Geir om aa aapne Prisoversikten i Excel igjen. I-kolonnen (Beregnet pris)
beregnes automatisk basert paa de nye H-verdiene.

## Spesialtilfeller aa kontrollere

| Linje | Kunde             | OBS                                                          |
|------:|-------------------|--------------------------------------------------------------|
|    28 | Veidekke e-post   | H-kolonnen = antall e-poster (manuelt) – IKKE fra Zisson    |
|    65 | Enova Trondheim   | Kvartalsfakturering – kontroller om det er faktureringmnd   |
|    67 | NYD               | Neste faktura august – kontroller datoer                     |
|    68 | Feste Trafikkskole| Faktureres september                                         |
| 50–54 | NEMUS-klinikker   | Manuell fordeling via NEMUS_Anrop_Fordeling-filen            |

## Ukjente koer i outputen

Hvis scriptet rapporterer ukjente koer, betyr det at en ny kunde er lagt til
i Zisson men ikke finnes i mappingen i `oppdater_h_kolonne.py`. I saa fall:

1. Finn riktig linjenummer i Prisoversikten for den nye kunden
2. Legg til en ny linje i `KO_LINJE`-lista i scriptet:
   ```python
   ("Koenavn-substring",   linje_nummer),
   ```
3. Kjoer torkjoring paa nytt for aa verifisere

## Ny maanesrapport med annet filnavn

Scriptet finner automatisk nyeste fil som matcher `We4you* M*ned*.xlsx`.
Hvis filen har et annet navn, oppgi det eksplisitt:

```bash
python3 "C:\Users\hei\Claude\Projects\Fakturaplan\oppdater_h_kolonne.py" "filnavn.xlsx" --skriv
```
