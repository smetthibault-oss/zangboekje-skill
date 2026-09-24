---
name: zangboekje-lokaal
description: Maak een Zie Ze Zingen-zangboekje (.docx) rechtstreeks in Claude, zonder naar app.thibaultsmet.be te moeten gaan. Gebruik deze skill wanneer de gebruiker vraagt om een zangboekje te maken, liedjes te bundelen, of een boekje te genereren voor een locatie/datum. Werkt met dezelfde huisstijl (Anton-titel, QR's, logo) als de webapp.
---

# Zangboekje maken (lokaal, zonder de webapp)

Deze skill genereert exact hetzelfde `.docx`-zangboekje als app.thibaultsmet.be
(zelfde titelblok, voetteksten, refrein-opmaak), maar volledig lokaal: jij haalt
de liedteksten op via je eigen Google Drive-connector, en een bijgeleverd
Node-scriptje (`generate.js`) bouwt er het Word-document van.

**Vereisten bij de gebruiker:**
- Een Google Drive-MCP-connector die verbonden is met een account dat toegang
  heeft tot de Zie Ze Zingen-liedjesmappen (zelfde Drive als de webapp gebruikt).
- Node.js, en éénmalig `npm install` uitgevoerd in deze skill-map.

Als er geen Drive-connector beschikbaar is in de huidige sessie, zeg dat
expliciet en vraag de gebruiker die eerst te verbinden (of geef aan dat je de
liedjesteksten ook rechtstreeks geplakt kan krijgen als alternatief).

## Vaste configuratie

- **Liedjesmappen** (doorzoek beide; de tweede is een aparte map met losse/
  nieuwere liedjesteksten die niet in de hoofdmap staan):
  - Hoofdmap: `1GbeENs6xVQSepbiw20Hj59voZi6POvXL`
  - Extra map ("APARTE LIEDJESTEKSTEN"): `1f6igPqbcTX_SAMsBfGsrWLOuC1DQhOTp`
- **Bestandsnaam van het resultaat**: `JJJJ-MM-DD Locatie.docx`

## Workflow

### Stap 0 — Skill bijwerken (stil, elke keer)

Deze map is een git-clone van https://github.com/smetthibault-oss/zangboekje-skill.
Haal vóór je begint de laatste versie op, zodat aanpassingen van Thibault
(opmaak, teksten, fixes) automatisch meekomen:

```bash
git pull --ff-only
```

(Working directory = deze skill-map.) Als er iets nieuws binnenkwam in
`package.json`, draai dan ook `npm install`. Als `git pull` faalt (geen
internet, of de map is geen git-clone omdat de skill uit een zip kwam), ga dan
gewoon verder met de lokale versie en zeg er kort bij dat de skill mogelijk niet
up-to-date is — blokkeer de rest van de workflow er niet op.

### Stap 1 — Info verzamelen

Vraag (in één bericht als het nog ontbreekt): **locatie** + **datum** + de
**lijst liedjes** (titel + artiest, of een genummerde lijst zoals gebruikers
die ook op de webapp plakken, bv. "1. Titel - ARTIEST").

### Stap 2 — Liedjes opzoeken in Drive

Gebruik de Google Drive-connector die in deze sessie beschikbaar is (het
zoek-tool, vaak `search_files` genoemd) om elke titel te vinden. Zoek in
**beide** mappen hierboven (`parentId` van elke map, met `or` ertussen als de
tool dat toelaat, anders twee losse zoekopdrachten). Voeg geen exacte
schrijfwijze-eisen toe — zoek los op een paar kernwoorden uit de titel, want
brongebruikers spellen titels niet altijd identiek (streepjes, hoofdletters,
"Sera" vs "Sera Sera", ...).

Kies bij twijfel tussen meerdere treffers de meest recent aangepaste versie,
en negeer bestanden die beginnen met `0_` of met `Kopie van ` (werkkopieën).

Meld liedjes die je nergens vindt expliciet aan de gebruiker in plaats van ze
zomaar over te slaan.

### Stap 3 — Tekst ophalen (parallel, één beurt)

Haal de volledige inhoud van elk gevonden document op (het lees-tool, vaak
`read_file_content` genoemd), voor alle gevonden liedjes in dezelfde
berichtbeurt. De inhoud komt terug als platte tekst waarin een origineel vet
gedrukte regel als `**zo**` genoteerd staat — laat die notatie exact zo staan,
verander er niets aan (het generatiescript herkent dat zelf).

### Stap 4 — Payload schrijven

Schrijf (Write-tool) een JSON-bestand, bv. naar
`<skill-map>/tmp/payload.json`:

```json
{
  "locatie": "WZC Ter Hovingen",
  "datum": "2026-10-22",
  "songs": [
    { "naam": "Titel - ARTIEST", "lyrics": "eerste regel\ntweede regel\n\n**Refrein-regel**\n..." }
  ]
}
```

- `naam`: exact zoals de gebruiker die opgaf (titel + artiest).
- `lyrics`: de opgehaalde tekst. Verwijder de titel-regel niet zelf — dat doet
  het script (het herkent en negeert een eerste regel die op de titel lijkt).
- Een lege regel = een stanza/coupletbreuk. Een regel die uitsluitend
  "Refrein", "Chorus", "Ref" (met of zonder ":") is, wordt automatisch **vet**
  "REFREIN" — typ in dat geval de volledige refreintekst dus niet nog eens uit.

### Stap 5 — Genereren

Voer uit (Bash-tool, working directory = deze skill-map):

```bash
node generate.js tmp/payload.json "tmp/2026-10-22 WZC Ter Hovingen.docx"
```

Bij de eerste keer op een machine: als dit faalt met `Cannot find module 'docx'`,
draai eerst `npm install` in deze map en probeer opnieuw.

### Stap 6 — Opleveren

Bevestig het resultaat (bestandsgrootte + aantal liedjes staat in de output)
en geef het `.docx`-bestand aan de gebruiker (SendUserFile of gelijkaardig,
afhankelijk van de omgeving).

## Kwaliteitscontrole

- De cover toont de volledige liedjeslijst.
- Geen liedje heeft een dubbele titelregel (eenmaal als kop, eenmaal als eerste
  tekstregel) — als dat toch gebeurt, wijkt de eerste regel van de brontekst te
  veel af van de opgegeven `naam`; pas dan zo nodig de `naam` aan zodat hij dichter
  bij de brontekst-titel ligt.
- Een herhaald refrein staat er als kort **REFREIN**, niet als de volledige tekst.
- Elke liedjespagina heeft een paginanummer links, logo in het midden, QR rechts
  in de voettekst; de eerste en laatste pagina tonen het logo + contactgegevens
  gecentreerd onderaan.
