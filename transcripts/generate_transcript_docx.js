// Builds a formatted Word document from the raw Groceryshop closing-panel transcript.
// The raw text is split at the start phrases below, so the transcript stays verbatim.
// Usage: npm install docx && node generate_transcript_docx.js
const fs = require("fs");
const path = require("path");
const {
  Document, Packer, Paragraph, TextRun, HeadingLevel, AlignmentType, Table, TableRow,
  TableCell, WidthType, ShadingType, BorderStyle, Footer, PageNumber, LevelFormat,
} = require("docx");

const RAW = fs.readFileSync(path.join(__dirname, "groceryshop_closing_panel_raw.txt"), "utf8").trim();
const OUT = path.join(__dirname, "Groceryshop_Closing_Panel_Transcript.docx");

const BEN = "Ben (Host / Moderator)";
const DEB = "Deborah Weinswig";
const KEL = "Kelly Kennedy";
const CHR = "Chris Walton";
const PAN = "Panelist (unclear)";

const SPEAKER_COLORS = { [BEN]: "1F4E79", [DEB]: "7B2C83", [KEL]: "2E7D32", [CHR]: "C55A11", [PAN]: "595959" };

// Each turn starts at the given phrase. A `section` opens a new heading at that turn.
const TURNS = [
  { section: "Opening & Panel Introductions", speaker: BEN, start: "Um, firstly, for everybody in the room" },
  { section: "Q1. Agentic Commerce: Is there a clearer picture of the endgame for grocery?", speaker: BEN, start: "Our 1st question is on agentic" },
  { speaker: CHR, start: "I'm first. I'm afraid to continue." },
  { speaker: BEN, start: "Interesting, interesting. What's that?" },
  { speaker: KEL, start: "Yeah, I'm going to a bad person here." },
  { speaker: BEN, start: "Sure thing. Denver, yeah" },
  { speaker: DEB, start: "I'm gonna tell you exactly where you've all headed." },
  { speaker: BEN, start: "Sure. Okay, I'm gonna, I'm gonna give you" },
  { speaker: CHR, start: "I mean, I think the discant immediation point" },
  { speaker: BEN, start: "Yeah, I think that, and therefore, driving the creation" },
  { section: "Q2. “Show Me the Money”: The best AI use case of the week", speaker: BEN, start: "Okay, we talked a lot about AI." },
  { speaker: KEL, start: "Okay, awesome. And I'm going to steal" },
  { speaker: BEN, start: "Chris, what about yourself? AI use case." },
  { speaker: CHR, start: "Oh, man. Well, you know, I'm vending" },
  { speaker: BEN, start: "Really? Yeah, great, good. Denver" },
  { speaker: DEB, start: "I think we heard it from the stage with Instacart CEO" },
  { speaker: BEN, start: "Right. I've got an example." },
  { speaker: PAN, start: "The thing wasn't strawberries." },
  { speaker: BEN, start: "Yes, yeah, that would be a long harder" },
  { section: "Q3. “How Soon Is Now?”: Shoppers’ need for speed", speaker: BEN, start: "Okay, how soon is now?" },
  { speaker: DEB, start: "Thank you for letting me take the lead." },
  { speaker: BEN, start: "Chris, are you thinking need to speed" },
  { speaker: CHR, start: "And everything, yeah." },
  { speaker: BEN, start: "You see, I'm thinking about multi places" },
  { speaker: DEB, start: "I have to jump in because..." },
  { speaker: BEN, start: "Okay, well, we're always out of time, but Chris" },
  { speaker: CHR, start: "That's where I come back to this intermediation" },
  { speaker: BEN, start: "I'm gonna... I'll try and get it in the next car." },
  { section: "Q4. “Supermarket Sweep”: How supermarket formats can respond to losing share", speaker: BEN, start: "The next question I'm okay, Jimmy." },
  { speaker: DEB, start: "Yes, because Christmas set me up perfect." },
  { speaker: BEN, start: "Okay, we can bring so many friends together." },
  { speaker: KEL, start: "Yeah, so what I'm gonna say is gonna sound" },
  { speaker: BEN, start: "Interesting, interesting. Chris, what have you heard" },
  { speaker: CHR, start: "Yeah, I mean, I think I think the interesting question" },
  { speaker: BEN, start: "Yeah. And so this was the exact premise" },
  { section: "Q5. Tech Talks: What technology excited the panel this week?", speaker: BEN, start: "All right, tech talks." },
  { speaker: DEB, start: "So there were two winners" },
  { speaker: BEN, start: "Kenny, we're excited to do tech wise" },
  { speaker: KEL, start: "Yeah, so I'm also going to talk about startup pitch" },
  { speaker: BEN, start: "Because I'm slightly concerned about asking you this." },
  { speaker: CHR, start: "Oh, really?" },
  { speaker: BEN, start: "Has anything excited you this week?" },
  { speaker: CHR, start: "Wow. No, from a tech perspective." },
  { speaker: BEN, start: "Yeah, and I think we had up Office of Wizster" },
  { section: "Q6. Stat of the Show", speaker: BEN, start: "Okay, we've got 10 minutes left" },
  { speaker: CHR, start: "Have the show. That one for me" },
  { speaker: BEN, start: "Yeah, I think that's very fat." },
  { speaker: KEL, start: "Yeah, so mine is from a Kinsey's" },
  { speaker: BEN, start: "Because 85%, okay, is big" },
  { speaker: DEB, start: "So, I've always been very focused on shop availability" },
  { speaker: BEN, start: "Okay. Brudy." },
  { section: "Q7. Quote of the Show", speaker: BEN, start: "All right, cool. Quote of the show" },
  { speaker: CHR, start: "Oh, man, okay, for the show." },
  { speaker: BEN, start: "I also really like Myron's" },
  { speaker: KEL, start: "So minus Ron Stewart" },
  { speaker: BEN, start: "Amen. Deborah, quote to the show." },
  { speaker: DEB, start: "Shoppers not asking for AI." },
  { section: "Closing Remarks", speaker: BEN, start: "I'm going to wrap with a quote" },
];

// Split raw text into turns; fail loudly if a phrase is missing or out of order.
function splitTurns() {
  const idx = TURNS.map((t) => {
    const i = RAW.indexOf(t.start);
    if (i < 0) throw new Error(`Start phrase not found: ${t.start}`);
    return i;
  });
  idx.forEach((v, k) => { if (k && v <= idx[k - 1]) throw new Error(`Out of order: ${TURNS[k].start}`); });
  if (idx[0] !== 0) throw new Error("First turn must start at beginning of transcript");
  const turns = TURNS.map((t, k) => ({ ...t, text: RAW.slice(idx[k], idx[k + 1] ?? RAW.length).trim() }));
  const rejoined = turns.map((t) => t.text).join(" ").replace(/\s+/g, " ");
  if (rejoined !== RAW.replace(/\s+/g, " ")) throw new Error("Split transcript does not match raw text");
  return turns;
}

// ---------- layout helpers ----------
const FONT = "Calibri";
const CONTENT_W = 9360; // US Letter with 1" margins
const border = { style: BorderStyle.SINGLE, size: 4, color: "BFBFBF" };
const borders = { top: border, bottom: border, left: border, right: border };

const p = (text, opts = {}) => new Paragraph({ spacing: { after: 120 }, ...opts, children: [new TextRun({ text, ...(opts.run || {}) })] });
const h1 = (text) => new Paragraph({ heading: HeadingLevel.HEADING_1, children: [new TextRun(text)] });
const h2 = (text) => new Paragraph({ heading: HeadingLevel.HEADING_2, children: [new TextRun(text)] });
const bullet = (runs) => new Paragraph({ numbering: { reference: "bullets", level: 0 }, spacing: { after: 60 }, children: runs });

function cell(text, width, { header = false, fill } = {}) {
  return new TableCell({
    borders,
    width: { size: width, type: WidthType.DXA },
    shading: fill ? { fill, type: ShadingType.CLEAR, color: "auto" } : undefined,
    margins: { top: 80, bottom: 80, left: 120, right: 120 },
    children: [new Paragraph({ children: [new TextRun({ text, bold: header, color: header ? "FFFFFF" : undefined, size: 20 })] })],
  });
}

function table(widths, header, rows) {
  return new Table({
    width: { size: CONTENT_W, type: WidthType.DXA },
    columnWidths: widths,
    rows: [
      new TableRow({ tableHeader: true, children: header.map((h, i) => cell(h, widths[i], { header: true, fill: "1F4E79" })) }),
      ...rows.map((r, ri) => new TableRow({ children: r.map((c, i) => cell(c, widths[i], { fill: ri % 2 ? "F2F2F2" : undefined })) })),
    ],
  });
}

// ---------- document content ----------
function buildChildren(turns) {
  const children = [];

  // Title block
  children.push(
    new Paragraph({ alignment: AlignmentType.CENTER, spacing: { before: 1200, after: 120 }, children: [new TextRun({ text: "GROCERYSHOP", bold: true, size: 28, color: "7F7F7F", characterSpacing: 60 })] }),
    new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 120 }, children: [new TextRun({ text: "Closing Session: Show Wrap-Up Panel", bold: true, size: 48, color: "1F4E79" })] }),
    new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 480 }, border: { bottom: { style: BorderStyle.SINGLE, size: 12, color: "1F4E79", space: 8 } }, children: [new TextRun({ text: "Session Transcript", size: 28, italics: true, color: "404040" })] }),
  );

  // Session overview
  children.push(h1("Session Overview"));
  children.push(table([2600, 6760], ["Item", "Details"], [
    ["Event", "Groceryshop (Las Vegas)"],
    ["Session", "Final session of the show: closing wrap-up panel"],
    ["Format", "Moderated panel; recorded for release as a podcast episode"],
    ["Session topic", "What the panel learned at the show: the good, the bad and “hopefully not too ugly”, answered through a set of themed questions"],
    ["Questions covered", "7 of the 8 planned questions (time ran out)"],
    ["Transcript", "Auto-generated, reproduced verbatim below (see Transcription Notes)"],
  ]));

  children.push(h1("Speakers"));
  children.push(table([2400, 3000, 3960], ["Speaker", "Role", "Organisation / Notes"], [
    ["Ben", "Host / Moderator", "Surname not stated in the recording; addressed as “Ben” by the panel"],
    ["Deborah Weinswig", "Panelist", "Founder & CEO, Coresight Research (transcribed as “Deborah Wyndstick”, “Causelight Research”); runs the Groceryshop startup pitch competition"],
    ["Kelly Kennedy", "Panelist", "Content Director, Groceryshop; responsible for the show’s on-stage content (name as heard in the recording)"],
    ["Chris Walton", "Panelist", "President & CEO (as introduced), Omni Talk Retail, “the media network for retailers, run by retailers”"],
  ]));

  children.push(h1("Session Agenda"));
  turns.filter((t) => t.section).forEach((t) => children.push(bullet([new TextRun(t.section)])));

  children.push(h1("Key Takeaways"));
  [
    ["Agentic commerce: ", "No agreed endgame yet. Disintermediation was the host’s word of the show; retailers are building agentic capabilities on their own assets, and the panel flagged the idea that “the LLM becomes the store.”"],
    ["Best AI use cases: ", "Sam’s Club member product testing, Albert Heijn’s “problem-first” associate assistant, Instacart’s AI-assisted baskets (reported 50% larger), and a premium blueberry grower using satellite and drone imagery."],
    ["Need for speed: ", "Speed of decision and speed of receipt both matter. Live streaming, AI that cuts cognitive load, and retailers joining several delivery marketplaces were all cited."],
    ["Supermarket formats: ", "Private label, strong digital storefronts, price investment and transparency, loyalty, and a single clear banner identity (the Hy-Vee example)."],
    ["Technology: ", "Practical applications over flashy demos: shelf-life / expiry tools, Perfect Store Group, Blue Collar Robotics centre-store picking, and an agentic media planner."],
    ["Stats of the show: ", "Instacart grocers with in-store pricing growing 10 points faster; 85% of consumers rate private label as equal or better; 35% better on-shelf availability with computer vision and robots; 41% year-on-year gas price rise; 8% of Albertsons trips come from forgotten ingredients."],
    ["Quotes of the show: ", "“The item sells itself”; “Rethink what innovation is. It has to start with the consumer”; “Shoppers are not asking for AI, they are asking for less work”; agents should be proactive and act, not wait to be searched."],
  ].forEach(([b, t]) => children.push(bullet([new TextRun({ text: b, bold: true }), new TextRun(t)])));

  children.push(h1("Transcription Notes"));
  children.push(p("The transcript below is the automatic speech-to-text output, kept word for word. Speaker labels were added from context (who was addressed and who replied) and are best-effort. Many names and terms were mis-heard by the transcription. Likely intended terms are listed below for reference only; they have not been changed in the transcript text.", { run: { size: 20 } }));
  children.push(table([3400, 5960], ["As transcribed", "Likely intended"], [
    ["Deborah Wyndstick / Causelight Research", "Deborah Weinswig / Coresight Research"],
    ["grocery shop / Grotop", "Groceryshop"],
    ["Omnitalk / Army talk", "Omni Talk (Retail)"],
    ["agentic cock / adjective / a gentic / a genti", "agentic commerce / agentic"],
    ["is Cart CEO / Chris Rogers", "Instacart CEO Chris Rogers"],
    ["Magloo and Brazil / Lou", "Magalu (Brazil) / Lu, its virtual influencer"],
    ["Schnuts", "Schnucks"],
    ["Albert Hyne", "Albert Heijn"],
    ["Hybee / Hiving / high-deep", "Hy-Vee"],
    ["Jerry Gosh / Jeremy Gosher", "Jeremy Gosch (Hy-Vee)"],
    ["Kinsey's", "McKinsey"],
    ["Alkins Media Connective", "Albertsons Media Collective"],
    ["Denver / Brudy (when addressing the panel)", "Deborah"],
    ["Kenny (when addressing the panel)", "Kelly"],
    ["Christmas (e.g. “Christmas point”)", "Chris’s (e.g. “Chris’s point”)"],
  ]));

  // Transcript
  children.push(new Paragraph({ pageBreakBefore: true, heading: HeadingLevel.HEADING_1, children: [new TextRun("Full Transcript")] }));
  for (const t of turns) {
    if (t.section) children.push(h2(t.section));
    children.push(new Paragraph({
      spacing: { before: 120, after: 40 },
      keepNext: true,
      children: [new TextRun({ text: t.speaker.toUpperCase(), bold: true, size: 20, color: SPEAKER_COLORS[t.speaker], characterSpacing: 20 })],
    }));
    children.push(new Paragraph({
      spacing: { after: 160, line: 300 },
      indent: { left: 360 },
      border: { left: { style: BorderStyle.SINGLE, size: 12, color: SPEAKER_COLORS[t.speaker], space: 10 } },
      children: [new TextRun({ text: t.text, size: 21 })],
    }));
  }
  return children;
}

const turns = splitTurns();
const doc = new Document({
  creator: "Groceryshop transcript",
  title: "Groceryshop Closing Session Transcript",
  styles: {
    default: { document: { run: { font: FONT, size: 22 } } },
    paragraphStyles: [
      { id: "Heading1", name: "Heading 1", basedOn: "Normal", next: "Normal", quickFormat: true,
        run: { size: 32, bold: true, font: FONT, color: "1F4E79" },
        paragraph: { spacing: { before: 360, after: 160 }, outlineLevel: 0 } },
      { id: "Heading2", name: "Heading 2", basedOn: "Normal", next: "Normal", quickFormat: true,
        run: { size: 26, bold: true, font: FONT, color: "2E75B6" },
        paragraph: { spacing: { before: 320, after: 120 }, outlineLevel: 1,
          border: { bottom: { style: BorderStyle.SINGLE, size: 4, color: "BDD7EE", space: 4 } } } },
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
        new TextRun({ text: "Groceryshop Closing Session Transcript  |  Page ", size: 16, color: "7F7F7F" }),
        new TextRun({ children: [PageNumber.CURRENT], size: 16, color: "7F7F7F" }),
      ] })] }),
    },
    children: buildChildren(turns),
  }],
});

Packer.toBuffer(doc).then((buf) => {
  fs.writeFileSync(OUT, buf);
  console.log(`Wrote ${OUT} (${turns.length} speaker turns)`);
});
