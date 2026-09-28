// Zuivere tekst-hulpfuncties, letterlijk overgenomen uit de webapp se songData.js
// (../../../songData.js) — geen Google-afhankelijkheden, dus hier los
// bruikbaar zonder OAuth/Drive-API.

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

// " - " (of een en/em-streepje) scheidt doorgaans titel van artiest, zowel in
// de opgegeven naam als in de eerste regel van de tekst zelf.
const DASH_SPLIT = /\s[-–—]\s/;
function titleOnly(text) {
  return normalizeForCompare((text || '').split(DASH_SPLIT)[0]);
}

// Sommige liedteksten herhalen de titel als eerste regel — die willen we
// overslaan, want de titel wordt al apart als kop boven het liedje gezet.
function stripLeadingTitleLine(naam, lines) {
  if (!lines.length) return lines;

  const first = lines[0];
  const isHeadingStyle = first.style?.startsWith('HEADING');

  const firstFull = normalizeForCompare(first.text);
  const firstTitle = titleOnly(first.text);
  const naamFull = normalizeForCompare(naam);
  const naamTitle = titleOnly(naam);

  // Enkel EXACTE overeenkomst (na normaliseren) — een "bevat"-check knipte
  // soms de eerste ECHTE gezongen regel weg zodra die toevallig de
  // titelwoorden bevatte. Zonder apart title-regel in de brontekst blijft de
  // eerste regel daardoor voortaan altijd staan.
  const exact = (a, b) => a.length > 0 && a === b;
  const looksLikeTitle = exact(firstFull, naamFull) || exact(firstFull, naamTitle)
    || exact(firstTitle, naamTitle) || exact(firstTitle, naamFull);

  return (isHeadingStyle || looksLikeTitle) ? lines.slice(1) : lines;
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

// Zet platte, geplakte liedjestekst om naar dezelfde { text, style, para }-vorm
// die de webapp gebruikt voor Google Docs-content. Een regel die volledig
// **vet** staat (of tussen [blokjes]) wordt als REFREIN/kop herkend; de rest
// is normale tekst. Een soft line break (Shift+Enter) geef je door als \v.
function lyricsToLines(lyrics) {
  return sanitizeText(lyrics).replace(/\r\n/g, '\n').split('\n').map((rawText) => {
    const boldMatch = rawText.trim().match(/^\*\*(.+)\*\*$/);
    const text = boldMatch ? boldMatch[1].trim() : rawText;
    const isHeading = /^\[.*\]$/.test(rawText.trim());
    return {
      text,
      style: isHeading ? 'HEADING_2' : 'NORMAL_TEXT',
      para: { elements: [{ textRun: { content: text, textStyle: { bold: !!boldMatch } } }] },
    };
  });
}

module.exports = {
  sanitizeText, normalizeRefrein, stripLeadingTitleLine, getLineStyle, lyricsToLines,
};
