// Zuivere tekst-hulpfuncties, grotendeels overgenomen uit de webapp (songData.js)
// — geen Google-afhankelijkheden, dus hier los bruikbaar zonder OAuth/Drive-API.

// Andere controletekens dan \v (soft line break) zijn ongeldig in Word-XML
// en breken het document, dus die vervangen we door een spatie.
function sanitizeText(text) {
  return (text || '').replace(/[\x00-\x08\x0C\x0E-\x1F]/g, ' ');
}

// Herkent een "refrein"-regel (Refrein, refrein:, [Refrein], Chorus, bis, ...) ongeacht
// hoofd-/kleine letters, en normaliseert hem naar bv. "REFREIN".
function normalizeRefrein(text) {
  const trimmed = text.trim();
  const isRefrein = /^[\[(]?\s*(\d+\s*x\s*)?(refrein|chorus)(\s*\d+\s*x)?(\s*\(?bis\)?)?\s*[\])]?\s*[:.…]*\s*$/i.test(trimmed);
  return isRefrein ? trimmed.toUpperCase() : null;
}

function normalizeForCompare(s) {
  return (s || '').trim().toLowerCase().replace(/[^\p{L}\p{N}]+/gu, ' ').trim();
}

// " - " (of een en/em-streepje) scheidt doorgaans titel van artiest.
const DASH_SPLIT = /\s[-–—]\s/;
function titleOnly(text) {
  return normalizeForCompare((text || '').split(DASH_SPLIT)[0]);
}

// Sommige liedteksten herhalen de titel als eerste regel — die willen we
// overslaan, want de titel wordt al apart als kop boven het liedje gezet.
// Verwijderd worden: alle opeenvolgende "titel"-regels bovenaan (markdown-kop
// "# ..."), of een eerste regel die EXACT (na normaliseren) de titel is.
// Een eerste regel die gewoon tekst is blijft altijd staan.
function stripLeadingTitleLine(naam, lines) {
  if (!lines.length) return lines;

  let start = 0;
  while (start < lines.length && lines[start].style === 'HEADING_1') start++;

  if (start === 0) {
    const first = lines[0];
    const naamFull = normalizeForCompare(naam);
    const naamTitle = titleOnly(naam);
    const exact = (x, y) => x.length > 0 && x === y;
    if (exact(normalizeForCompare(first.text), naamFull) || exact(normalizeForCompare(first.text), naamTitle)
      || exact(titleOnly(first.text), naamTitle) || exact(titleOnly(first.text), naamFull)) {
      start = 1;
    }
  }

  while (start > 0 && start < lines.length && lines[start].blank) start++;
  return lines.slice(start);
}

// Geeft de opmaak (vet/cursief) van een alinea terug, gebaseerd op de eerste
// run die ECHT tekst bevat.
function getLineStyle(para) {
  const run = para.elements?.find(el => el.textRun?.content?.trim())?.textRun
    || para.elements?.[0]?.textRun;
  return {
    isBold: run?.textStyle?.bold || false,
    isItalic: run?.textStyle?.italic || false,
  };
}

// Leest één regel met markdown-sterretjes. "state" loopt door over regels heen
// (de Drive-connector zet een vette alinea soms over meerdere regels, met
// "****" op de grens = sluiten + weer openen). Een reeks van n sterretjes
// schakelt floor(n/2) keer vet en (n mod 2) keer cursief om.
function parseInline(line, state) {
  let out = '';
  let total = 0, boldChars = 0, italicChars = 0, last = 0, m;
  const add = (chunk) => {
    if (!chunk) return;
    out += chunk;
    const n = chunk.replace(/\s/g, '').length;
    total += n;
    if (state.bold) boldChars += n;
    if (state.italic) italicChars += n;
  };
  const re = /\*+/g;
  while ((m = re.exec(line))) {
    add(line.slice(last, m.index));
    const n = m[0].length;
    if (Math.floor(n / 2) % 2 === 1) state.bold = !state.bold;
    if (n % 2 === 1) state.italic = !state.italic;
    last = m.index + n;
  }
  add(line.slice(last));
  return {
    text: out.replace(/\\([\\`*_{}\[\]()#+\-.!])/g, '$1').trim(),
    bold: total > 0 && boldChars >= total / 2,
    italic: total > 0 && italicChars >= total / 2,
  };
}

function makeLine(text, style, bold, italic, blank) {
  return {
    text, style, blank: !!blank,
    para: { elements: [{ textRun: { content: text, textStyle: { bold, italic } } }] },
  };
}

// Zet de tekst uit Drive (of platte, geplakte tekst) om naar dezelfde
// { text, style, para }-vorm die de webapp voor Google Docs-content gebruikt.
//
// Twee soorten invoer:
//  - Drive-connector (markdown): elke alinea staat gescheiden door een lege regel,
//    een LEGE alinea (strofe-einde) is een regel met enkel spaties. Zachte
//    regeleinden binnen een alinea eindigen op 2 spaties.
//  - Platte tekst: één regel per regel, een lege regel = strofe-einde.
// "**vet**", "*cursief*", "# kop" en "[Label]" worden herkend.
function lyricsToLines(lyrics) {
  const raw = sanitizeText(lyrics).replace(/\r\n/g, '\n');
  const markdownExport = /^ +$/m.test(raw);
  const lines = [];
  let state = { bold: false, italic: false };

  const pushBlank = () => {
    if (lines.length && !lines[lines.length - 1].blank) lines.push(makeLine('', 'NORMAL_TEXT', false, false, true));
  };

  for (const rawLine of raw.split('\n')) {
    if (rawLine === '') {                       // alinea-scheiding
      state = { bold: false, italic: false };
      if (!markdownExport) pushBlank();
      continue;
    }
    if (!rawLine.trim()) { pushBlank(); continue; }   // lege alinea / lege regel

    const heading1 = /^\s*#{1,6}\s/.test(rawLine);
    const body = rawLine.replace(/^\s*#{1,6}\s+/, '');
    const { text, bold, italic } = parseInline(body, state);
    if (!text) continue;                        // enkel losse markeringen (bv. "**")

    const bracketLabel = /^\[.*\]$/.test(text);
    lines.push(makeLine(text, heading1 ? 'HEADING_1' : bracketLabel ? 'HEADING_2' : 'NORMAL_TEXT', bold, italic, false));
  }

  while (lines.length && lines[lines.length - 1].blank) lines.pop();
  while (lines.length && lines[0].blank) lines.shift();
  return lines;
}

module.exports = {
  sanitizeText, normalizeRefrein, stripLeadingTitleLine, getLineStyle, lyricsToLines,
};
