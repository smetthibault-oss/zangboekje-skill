# Zangboekje-skill (Zie Ze Zingen)

Genereert een Zie Ze Zingen-zangboekje (.docx) rechtstreeks in Claude, in
dezelfde huisstijl als app.thibaultsmet.be — zonder dat je naar de website moet
gaan. De skill werkt zichzelf bij: bij elk gebruik haalt Claude de laatste versie
van deze repo op.

## Installeren

Vereisten: Claude (Code of Desktop) met een **Google Drive-connector** die
toegang heeft tot de Zie Ze Zingen-liedjesmappen, en [Node.js](https://nodejs.org).

```bash
git clone https://github.com/smetthibault-oss/zangboekje-skill ~/.claude/skills/zangboekje-lokaal
cd ~/.claude/skills/zangboekje-lokaal
npm install
```

(Windows: `~` is `C:\Users\<jouw naam>`.)

Vraag daarna aan Claude bv.: *"Maak een zangboekje voor WZC Ter Hovingen op
22/10/2026 met: Imagine - John Lennon, Het dorp - Wim Sonneveld, ..."*

## Updaten

Gebeurt automatisch (stap 0 in `SKILL.md`). Handmatig kan ook:

```bash
cd ~/.claude/skills/zangboekje-lokaal && git pull && npm install
```

## Wat dit wel/niet is

- Maakt enkel het **.docx-zangboekje**, lokaal; er wordt niets geüpload.
- De PowerPoint-export, Genius-import en Google-login zitten enkel in de webapp.
- De opmaak (`generate.js`) is een zelfstandige kopie van de webapp-opmaak.
