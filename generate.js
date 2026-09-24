// Losse versie van ../../routes/generate.js: zelfde opmaak (blauwdruk-stijl),
// maar zonder Express/Google — de aanroeper (de Claude-skill) heeft de
// liedjesteksten al via de Drive-connector opgehaald en geeft ze als platte
// tekst mee. Zie SKILL.md voor de volledige workflow.
//
// Gebruik als CLI: node generate.js payload.json output.docx
// payload.json: { "locatie": "...", "datum": "JJJJ-MM-DD", "songs": [{ "naam": "...", "lyrics": "..." }] }

const fs = require('fs');
const path = require('path');
const { normalizeRefrein, stripLeadingTitleLine, getLineStyle, lyricsToLines } = require('./lib/songFormatting');

const {
  Document, Packer, Paragraph, TextRun, UnderlineType,
  ImageRun, Table, TableRow, TableCell, WidthType, BorderStyle, AlignmentType, Footer, PageNumber,
  CharacterSet, TextWrappingType, TextWrappingSide,
  HorizontalPositionRelativeFrom, HorizontalPositionAlign, VerticalPositionRelativeFrom,
} = require('docx');

const ASSETS = path.join(__dirname, 'assets');
const logoBuffer     = fs.readFileSync(path.join(ASSETS, 'docx-logo.png'));
const qrTitleBuffer  = fs.readFileSync(path.join(ASSETS, 'docx-qr-titel.png'));
const qrFooterBuffer = fs.readFileSync(path.join(ASSETS, 'docx-qr-voet.png'));
const antonBuffer    = fs.readFileSync(path.join(ASSETS, 'Anton-Regular.ttf'));

// ─── Hulpfuncties voor opmaak (identiek aan de webapp) ────────────────────────

function line(runs) {
  return new Paragraph({ spacing: { before: 0, after: 0, line: 240, lineRule: 'auto' }, children: runs });
}
function t(text, brk) { return new TextRun({ text, font: 'Calibri', size: 28, break: brk }); }
function bold(text, brk) { return new TextRun({ text, font: 'Calibri', size: 28, bold: true, break: brk }); }
function italic(text, brk) { return new TextRun({ text, font: 'Calibri', size: 28, italics: true, break: brk }); }
function boldItalic(text, brk) { return new TextRun({ text, font: 'Calibri', size: 28, bold: true, italics: true, break: brk }); }
function empty() { return line([t('')]); }

function songHeading(text, { pageBreakBefore = false } = {}) {
  return new Paragraph({
    pageBreakBefore,
    spacing: { before: 0, after: 120, line: 240, lineRule: 'auto' },
    children: [new TextRun({ text, bold: true, font: 'Calibri', size: 28, underline: { type: UnderlineType.SINGLE } })],
  });
}
function tocEntry(text, first) {
  return new Paragraph({
    spacing: { before: first ? 0 : 120, after: 0, line: 240, lineRule: 'auto' },
    children: [new TextRun({ text, font: 'Calibri', size: 28 })],
  });
}

function anton(text, size, opts = {}) {
  return new TextRun({ text, font: 'Anton', size, ...opts });
}

function titleBlock(subtitle, { pageBreakBefore = false } = {}) {
  return [
    new Paragraph({
      pageBreakBefore,
      spacing: { before: 0, after: 0, line: 276, lineRule: 'auto' },
      children: [
        anton('ZIE ZE ZINGEN!', 112, { bold: true }),
        new ImageRun({
          data: qrTitleBuffer,
          transformation: { width: 150, height: 190 },
          floating: {
            horizontalPosition: { relative: HorizontalPositionRelativeFrom.MARGIN, align: HorizontalPositionAlign.RIGHT },
            verticalPosition: { relative: VerticalPositionRelativeFrom.PAGE, offset: 747600 },
            wrap: { type: TextWrappingType.SQUARE, side: TextWrappingSide.BOTH_SIDES },
            allowOverlap: true,
            lockAnchor: true,
            behindDocument: false,
          },
        }),
      ],
    }),
    new Paragraph({ spacing: { before: 0, after: 0, line: 276, lineRule: 'auto' }, children: [anton(subtitle, 50)] }),
  ];
}

const noBorder = { style: BorderStyle.NONE, size: 0, color: 'FFFFFF' };
const noBorders = { top: noBorder, bottom: noBorder, left: noBorder, right: noBorder };
const TEXT_WIDTH = 8930;

function footerTable(columnWidths, cells, { alignment } = {}) {
  const width = columnWidths.reduce((a, b) => a + b, 0);
  return new Table({
    columnWidths,
    width: { size: width, type: WidthType.DXA },
    ...(alignment ? { alignment } : {}),
    layout: 'fixed',
    margins: { top: 0, bottom: 0, left: 0, right: 0 },
    borders: { ...noBorders, insideHorizontal: noBorder, insideVertical: noBorder },
    rows: [new TableRow({
      children: cells.map((children, i) => new TableCell({
        width: { size: columnWidths[i], type: WidthType.DXA }, borders: noBorders, verticalAlign: 'center', children,
      })),
    })],
  });
}

function pageNumberParagraph() {
  return new Paragraph({
    alignment: AlignmentType.LEFT,
    spacing: { before: 0, after: 0 },
    children: [new TextRun({ children: [PageNumber.CURRENT], font: 'Calibri', size: 28 })],
  });
}
function logoParagraph(px, alignment = AlignmentType.CENTER) {
  return new Paragraph({
    alignment,
    spacing: { before: 0, after: 0 },
    children: [new ImageRun({ data: logoBuffer, transformation: { width: px, height: px } })],
  });
}
function contactParagraphs(alignment = AlignmentType.LEFT) {
  const row = (label, value) => new Paragraph({
    alignment,
    spacing: { before: 0, after: 0, line: 276, lineRule: 'auto' },
    children: [anton(`${label}  ${value}`, 36)],
  });
  return [row('Boek ons :', 'info@ziezezingen.be'), row('Zing mee:', 'ziezezingen.be/events')];
}
// Logo en contactgegevens naast elkaar (zoals de blauwdruk), als geheel gecentreerd.
function logoAndContact(px) {
  return [footerTable([2800, 5430], [[logoParagraph(px)], contactParagraphs()])];
}
function firstPageFooter() { return new Footer({ children: logoAndContact(165) }); }
function songFooter() {
  return new Footer({
    children: [footerTable([2111, 4000, 2819], [
      [pageNumberParagraph()],
      [logoParagraph(115)],
      [new Paragraph({
        alignment: AlignmentType.RIGHT,
        spacing: { before: 0, after: 0 },
        children: [new ImageRun({ data: qrFooterBuffer, transformation: { width: 83, height: 100 } })],
      })],
    ])],
  });
}

function docLinesToParagraphs(lines) {
  const paragraphs = [];
  for (const { text, style, para } of lines) {
    if (!text.trim() && style === 'NORMAL_TEXT') { paragraphs.push(empty()); continue; }

    const refreinText = normalizeRefrein(text.replace(/\v/g, ' '));
    if (refreinText) { paragraphs.push(line([bold(refreinText)])); continue; }

    const { isBold, isItalic } = getLineStyle(para);

    if (style === 'HEADING_1' || style === 'HEADING_2') {
      paragraphs.push(songHeading(text.replace(/\v/g, ' ')));
    } else {
      const styleFn = isBold && isItalic ? boldItalic : isBold ? bold : isItalic ? italic : t;
      const segments = text.split('\v');
      const runs = segments.map((seg, i) => styleFn(seg, i > 0 ? 1 : undefined));
      paragraphs.push(line(runs));
    }
  }
  return paragraphs;
}

const WIE_ZIJN_WE_TEKST = [
  'Zie Ze Zingen is het samenzang collectief van Geertrui Coppens en Jasper De Mulder.',
  "Onze samenzang concerten zijn verbindende programmaties op festivals, in concertzalen en culturele centra. Met humor en een reeks uitgekiende zangtechnieken zorgen wij als 'zangopzwepers' voor magie: we transformeren elk publiek in een verbluffend, meerstemmig koor. Hierdoor ontstaat een wonderlijke samenhorigheid tussen volstrekt onbekenden. De lage instapdrempel en hoge wow-factor zorgen keer op keer voor kippenvel.",
  'Kunstencentrum VIERNULVIER in Gent is een van onze vaste partners. De afgelopen jaren gaven we concerten en workshops in een twintigtal culturele centra. Ook Trefpunt ziet ons graag jaarlijks terugkomen op de Gentse Feesten, waar honderden mensen uitkijken naar de speelse singalong avonden. In de donkere maanden brengen we dan weer troost met serene samenzangconcerten op begraafplaatsen, in concertzalen of ontwijde kerken.',
  'Daarnaast zetten we ons steeds vaker in met inclusieve projecten. Zo brengen we mensen samen via samenzang concerten met OKAN-klassen, volwassen anderstaligen, bewoners van woonzorgcentra en mensen met (jong)dementie.',
  'Bij Zie Ze Zingen koesteren we klein en groot. Massale samenzang geeft ons keer op keer een kick. Kleine singalongs schenken ons de voldoening van diepmenselijk contact.',
];

function wieZijnWe() {
  return [
    ...titleBlock('Wie zijn we?', { pageBreakBefore: true }),
    ...WIE_ZIJN_WE_TEKST.map((text, i) => new Paragraph({
      spacing: { before: i === 0 ? 240 : 0, after: 180, line: 276, lineRule: 'auto' },
      children: [t(text)],
    })),
    ...contactParagraphs(AlignmentType.CENTER),
  ];
}

// ─── Hoofdfunctie ──────────────────────────────────────────────────────────────

async function genereerZangboekje({ locatie, datum, songs }) {
  if (!locatie || !datum || !songs?.length) {
    throw new Error('locatie, datum en minstens 1 liedje zijn verplicht.');
  }

  const d = new Date(datum);
  const datumNL = Number.isNaN(d.getTime())
    ? datum
    : `${String(d.getDate()).padStart(2, '0')}.${String(d.getMonth() + 1).padStart(2, '0')}.${String(d.getFullYear()).slice(-2)}`;

  const songData = songs.map(s => ({ naam: s.naam, lines: lyricsToLines(s.lyrics || '') }));

  const children = [];
  children.push(...titleBlock(`${locatie} - ${datumNL}`), empty(), empty());
  songData.forEach((song, i) => children.push(tocEntry(`${i + 1}. ${song.naam}`, i === 0)));

  songData.forEach((song, i) => {
    children.push(songHeading(`${i + 1}. ${song.naam}`, { pageBreakBefore: true }));
    if (song.lines.length > 0) {
      const lines = stripLeadingTitleLine(song.naam, song.lines);
      children.push(...docLinesToParagraphs(lines));
    }
  });

  children.push(...wieZijnWe());

  const pageProps = {
    page: {
      size: { width: 11906, height: 16838 },
      margin: { top: 1417, right: 1134, bottom: 1134, left: 1842, footer: 200 },
    },
  };

  const doc = new Document({
    fonts: [{ name: 'Anton', data: antonBuffer, characterSet: CharacterSet.ANSI }],
    styles: { default: { document: { run: { font: 'Calibri', size: 24 } } } },
    sections: [
      {
        properties: { ...pageProps, titlePage: true },
        footers: { first: firstPageFooter(), default: songFooter() },
        children,
      },
    ],
  });

  return Packer.toBuffer(doc);
}

module.exports = { genereerZangboekje };

// ─── CLI ────────────────────────────────────────────────────────────────────

if (require.main === module) {
  const [payloadPath, outputPath, uploadArg] = process.argv.slice(2);
  if (!payloadPath || !outputPath) {
    console.error('Gebruik: node generate.js payload.json output.docx');
    process.exit(1);
  }
  const payload = JSON.parse(fs.readFileSync(payloadPath, 'utf-8'));
  genereerZangboekje(payload)
    .then(buffer => {
      fs.writeFileSync(outputPath, buffer);
      console.log(`OK: ${outputPath} (${Math.round(buffer.length / 1024)} KB, ${payload.songs.length} liedjes)`);

      // Optioneel: kopieer het resultaat naar een map die door Google Drive voor
      // desktop gesynchroniseerd wordt = "direct uploaden naar Drive".
      // Map via 3e argument, env ZANGBOEKJE_UPLOADMAP of het bestand uploadmap.txt.
      let uploadMap = uploadArg || process.env.ZANGBOEKJE_UPLOADMAP;
      const cfg = path.join(__dirname, 'uploadmap.txt');
      if (!uploadMap && fs.existsSync(cfg)) uploadMap = fs.readFileSync(cfg, 'utf-8').trim();
      if (uploadMap) {
        if (!fs.existsSync(uploadMap)) {
          console.error(`WAARSCHUWING: uploadmap bestaat niet: ${uploadMap} (niet gekopieerd)`);
        } else {
          const target = path.join(uploadMap, path.basename(outputPath));
          fs.copyFileSync(outputPath, target);
          console.log(`GEKOPIEERD NAAR DRIVE-MAP: ${target}`);
        }
      }
    })
    .catch(err => {
      console.error('FOUT:', err.message);
      process.exit(1);
    });
}
