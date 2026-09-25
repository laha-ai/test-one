// Builds one master Word document covering every Groceryshop session in sessions.json.
// Sessions with a transcript get their verbatim text split into speaker turns using the
// start phrases in the matching *_turns.json file; the rest are listed as "Transcript to follow".
// Usage: npm install docx && node generate_session_notes.js
const fs = require("fs");
const path = require("path");
const {
  Document, Packer, Paragraph, TextRun, HeadingLevel, AlignmentType, Table, TableRow,
  TableCell, WidthType, ShadingType, BorderStyle, Footer, PageNumber, LevelFormat,
} = require("docx");

const DIR = __dirname;
const OUT = path.join(DIR, "Groceryshop_2026_Session_Notes.docx");
const sessions = JSON.parse(fs.readFileSync(path.join(DIR, "sessions.json"), "utf8"));

// ---------- data helpers ----------
const MONTHS = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"];
const DAYS = ["Sun", "Mon", "Tue", "Wed", "Thu", "Fri", "Sat"];

function fmtDate(iso) {
  if (!iso) return "TBC";
  const [y, m, d] = iso.split("-").map(Number);
  const dt = new Date(Date.UTC(y, m - 1, d));
  return `${DAYS[dt.getUTCDay()]}, ${d} ${MONTHS[m - 1]} ${y}`;
}

function fmtTime(hhmm) {
  const [h, m] = hhmm.split(":").map(Number);
  return `${h % 12 || 12}:${String(m).padStart(2, "0")} ${h < 12 ? "AM" : "PM"}`;
}

const fmtSlot = (s) => (s.start ? `${fmtTime(s.start)} – ${fmtTime(s.end)} PT` : "TBC");

// Sessions in time order; sessions without a date go last, in file order.
function sortSessions(list) {
  return list
    .map((s, i) => ({ s, i }))
    .sort((a, b) => {
      const ka = a.s.date ? `${a.s.date} ${a.s.start}` : "~";
      const kb = b.s.date ? `${b.s.date} ${b.s.start}` : "~";
      return ka < kb ? -1 : ka > kb ? 1 : a.i - b.i;
    })
    .map((x) => x.s);
}

// Split a raw transcript at each turn's start phrase; fail loudly if the split is not lossless.
function loadTranscript(t) {
  const raw = fs.readFileSync(path.join(DIR, t.raw), "utf8").trim();
  const spec = JSON.parse(fs.readFileSync(path.join(DIR, t.turns), "utf8"));
  const idx = spec.turns.map((turn) => {
    const i = raw.indexOf(turn.start);
    if (i < 0) throw new Error(`${t.raw}: start phrase not found: ${turn.start}`);
    return i;
  });
  idx.forEach((v, k) => { if (k && v <= idx[k - 1]) throw new Error(`${t.raw}: out of order: ${spec.turns[k].start}`); });
  if (idx[0] !== 0) throw new Error(`${t.raw}: first turn must start at the beginning`);
  const turns = spec.turns.map((turn, k) => ({ ...turn, text: raw.slice(idx[k], idx[k + 1] ?? raw.length).trim() }));
  if (turns.map((x) => x.text).join(" ").replace(/\s+/g, " ") !== raw.replace(/\s+/g, " ")) {
    throw new Error(`${t.raw}: split transcript does not match raw text`);
  }
  return { ...spec, turns };
}

// ---------- layout helpers ----------
const FONT = "Calibri";
const NAVY = "1F4E79";
const CONTENT_W = 9360; // US Letter with 1" margins
const PALETTE = ["1F4E79", "7B2C83", "2E7D32", "C55A11", "00838F", "AD1457"];
const border = { style: BorderStyle.SINGLE, size: 4, color: "BFBFBF" };
const borders = { top: border, bottom: border, left: border, right: border };

const h = (level, text, opts = {}) => new Paragraph({ heading: level, ...opts, children: [new TextRun(text)] });
const p = (text, run = {}) => new Paragraph({ spacing: { after: 120 }, children: [new TextRun({ text, ...run })] });
const bullet = (runs) => new Paragraph({ numbering: { reference: "bullets", level: 0 }, spacing: { after: 60 }, children: runs });

function cell(text, width, { header = false, fill, bold = false } = {}) {
  return new TableCell({
    borders,
    width: { size: width, type: WidthType.DXA },
    shading: fill ? { fill, type: ShadingType.CLEAR, color: "auto" } : undefined,
    margins: { top: 80, bottom: 80, left: 120, right: 120 },
    children: [new Paragraph({ children: [new TextRun({ text, bold: header || bold, color: header ? "FFFFFF" : undefined, size: 20 })] })],
  });
}

function table(widths, header, rows, { boldFirstCol = false } = {}) {
  return new Table({
    width: { size: CONTENT_W, type: WidthType.DXA },
    columnWidths: widths,
    rows: [
      new TableRow({ tableHeader: true, children: header.map((t, i) => cell(t, widths[i], { header: true, fill: NAVY })) }),
      ...rows.map((r, ri) => new TableRow({
        cantSplit: true,
        children: r.map((c, i) => cell(c, widths[i], { fill: ri % 2 ? "F2F2F2" : undefined, bold: boldFirstCol && i === 0 })),
      })),
    ],
  });
}

function statusLabel(s) {
  return s.transcript ? "Transcript included" : "Transcript to follow";
}

// ---------- document sections ----------
function coverAndIndex(ordered) {
  const out = [
    new Paragraph({ alignment: AlignmentType.CENTER, spacing: { before: 1200, after: 120 }, children: [new TextRun({ text: "GROCERYSHOP 2026", bold: true, size: 28, color: "7F7F7F", characterSpacing: 60 })] }),
    new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 120 }, children: [new TextRun({ text: "Session Notes & Transcripts", bold: true, size: 48, color: NAVY })] }),
    new Paragraph({
      alignment: AlignmentType.CENTER, spacing: { after: 480 },
      border: { bottom: { style: BorderStyle.SINGLE, size: 12, color: NAVY, space: 8 } },
      children: [new TextRun({ text: "Sessions attended, speakers and transcripts", size: 26, italics: true, color: "404040" })],
    }),
    h(HeadingLevel.HEADING_1, "Sessions Attended"),
    table([1500, 1900, 1500, 3060, 1400], ["Date", "Time", "Stage", "Session", "Notes"],
      ordered.map((s) => [fmtDate(s.date), fmtSlot(s), s.stage || "TBC", s.title, statusLabel(s)])),
    p(""),
    p("Times are Los Angeles time (PT), as listed in the Groceryshop app. “TBC” means the detail has not been provided yet.", { size: 18, italics: true, color: "595959" }),
  ];
  return out;
}

function sessionSection(s, n, transcript) {
  const out = [];
  out.push(new Paragraph({ pageBreakBefore: true, spacing: { after: 60 }, children: [new TextRun({ text: `SESSION ${n}`, bold: true, size: 18, color: "7F7F7F", characterSpacing: 40 })] }));
  out.push(h(HeadingLevel.HEADING_1, s.title));

  const details = [
    ["Date", fmtDate(s.date)],
    ["Time", fmtSlot(s)],
    ["Stage", s.stage || "TBC"],
  ];
  if (s.sponsors?.length) details.push(["Sponsor", s.sponsors.join(", ")]);
  details.push(["Notes", statusLabel(s)]);
  if (s.title_note) details.push(["Title", s.title_note]);
  out.push(table([2200, 7160], ["Session Details", ""], details, { boldFirstCol: true }));

  out.push(h(HeadingLevel.HEADING_2, "Speakers"));
  const people = [
    ...(s.moderators || []).map((m) => [m.name, "Moderator", m.title, m.company]),
    ...(s.speakers || []).map((m) => [m.name, s.moderators?.length && s.speakers.length > 1 ? "Panelist" : "Speaker", m.title, m.company]),
  ];
  if (people.length) out.push(table([2200, 1400, 3160, 2600], ["Name", "Role", "Title", "Organisation"], people));
  else out.push(p("Speakers TBC.", { italics: true, color: "7F7F7F" }));
  if (s.speakers_note) out.push(p(s.speakers_note, { size: 18, italics: true, color: "595959" }));

  if (!transcript) {
    out.push(h(HeadingLevel.HEADING_2, "Transcript"));
    out.push(p("Transcript to follow.", { italics: true, color: "7F7F7F" }));
    return out;
  }

  if (transcript.takeaways?.length) {
    out.push(h(HeadingLevel.HEADING_2, "Key Takeaways"));
    transcript.takeaways.forEach(([b, t]) => out.push(bullet([new TextRun({ text: `${b}: `, bold: true }), new TextRun(t)])));
  }

  if (transcript.transcription_notes?.length) {
    out.push(h(HeadingLevel.HEADING_2, "Transcription Notes"));
    out.push(p("The transcript is the automatic speech-to-text output, kept word for word. Speaker labels were added from context (who was addressed and who replied). Names and terms the transcription misheard are listed below for reference; they have not been changed in the transcript text.", { size: 20 }));
    out.push(table([3400, 5960], ["As transcribed", "Likely intended"], transcript.transcription_notes));
  }

  // Colour per speaker, in order of first appearance.
  const colors = {};
  transcript.turns.forEach((t) => {
    if (!(t.speaker in colors)) colors[t.speaker] = t.speaker.includes("unclear") ? "595959" : PALETTE[Object.keys(colors).length % PALETTE.length];
  });

  out.push(h(HeadingLevel.HEADING_2, "Full Transcript", { pageBreakBefore: true }));
  for (const t of transcript.turns) {
    if (t.section) out.push(h(HeadingLevel.HEADING_3, t.section));
    out.push(new Paragraph({
      spacing: { before: 120, after: 40 }, keepNext: true,
      children: [new TextRun({ text: t.speaker.toUpperCase(), bold: true, size: 20, color: colors[t.speaker], characterSpacing: 20 })],
    }));
    out.push(new Paragraph({
      spacing: { after: 160, line: 300 }, indent: { left: 360 },
      border: { left: { style: BorderStyle.SINGLE, size: 12, color: colors[t.speaker], space: 10 } },
      children: [new TextRun({ text: t.text, size: 21 })],
    }));
  }
  return out;
}

// ---------- build ----------
const ordered = sortSessions(sessions);
const children = coverAndIndex(ordered);
let turnCount = 0;
ordered.forEach((s, i) => {
  const transcript = s.transcript ? loadTranscript(s.transcript) : null;
  if (transcript) turnCount += transcript.turns.length;
  children.push(...sessionSection(s, i + 1, transcript));
});

const heading = (id, name, size, color, before, extra = {}) => ({
  id, name, basedOn: "Normal", next: "Normal", quickFormat: true,
  run: { size, bold: true, font: FONT, color },
  paragraph: { spacing: { before, after: 140 }, keepNext: true, ...extra },
});

const doc = new Document({
  creator: "Groceryshop session notes",
  title: "Groceryshop 2026 Session Notes",
  styles: {
    default: { document: { run: { font: FONT, size: 22 } } },
    paragraphStyles: [
      heading("Heading1", "Heading 1", 34, NAVY, 240, { outlineLevel: 0 }),
      heading("Heading2", "Heading 2", 26, "2E75B6", 320, {
        outlineLevel: 1, border: { bottom: { style: BorderStyle.SINGLE, size: 4, color: "BDD7EE", space: 4 } },
      }),
      heading("Heading3", "Heading 3", 23, "404040", 280, { outlineLevel: 2 }),
    ],
  },
  numbering: {
    config: [{ reference: "bullets", levels: [{ level: 0, format: LevelFormat.BULLET, text: "•", alignment: AlignmentType.LEFT,
      style: { paragraph: { indent: { left: 720, hanging: 360 } } } }] }],
  },
  sections: [{
    properties: { page: { size: { width: 12240, height: 15840 }, margin: { top: 1440, right: 1440, bottom: 1440, left: 1440 } } },
    footers: {
      default: new Footer({ children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [
        new TextRun({ text: "Groceryshop 2026 Session Notes  |  Page ", size: 16, color: "7F7F7F" }),
        new TextRun({ children: [PageNumber.CURRENT], size: 16, color: "7F7F7F" }),
      ] })] }),
    },
    children,
  }],
});

Packer.toBuffer(doc).then((buf) => {
  fs.writeFileSync(OUT, buf);
  console.log(`Wrote ${OUT}: ${ordered.length} sessions, ${turnCount} transcript turns`);
});
