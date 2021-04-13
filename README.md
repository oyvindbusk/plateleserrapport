# Groovy app for å hente ut rapporter fra plateleser:
### Krever:
* Java må være installert
### Installering:
Kjør følgende kommando:
```sh
sh compile.sh
```
Dette lager mappa build, hvor programmet ligger i libs/app-1.0.jar. 

### Beskrivelse av bruk:
* Åpne program ved å dobbeltklikke på app-1.0.jar
* Åpne xlsc-fil fra plateleser
** Sample angis som "prøvenummer,år ekstraktID" eks: "900,45 55663"
* Output-fil lagres der programmet åpnes fra (altså build/libs/ hvis du åpner programmet derfra)
** Output er en fil med 8 kolonner, navn blir satt fra sessionID i resultatfil:
** Brønn, prøvenumer, år, ekstraktID, kons,	ratio, brukernavn, dagens dato

#### Gjør følgende:
- [x] Leser inn xlsx-fil
- [x] Henter ut kons og ratio
- [x] Legger til prøver i csv-fil

#### Beskrivelse av filer i mappe:
2020_03_05_1_hilt.xls -> Output fra plateleser
