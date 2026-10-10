// Printable English speech for two speakers (A4, black-and-white friendly)
const fs = require("fs");
const path = require("path");
const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell, WidthType, BorderStyle, ShadingType,
  AlignmentType, HeadingLevel, PageBreak, Footer, Header, PageNumber, VerticalAlign, LevelFormat, TableLayoutType,
} = require("docx");
const { SLIDES, WPM } = require("./notes3_en.js");
const EXTRA = require("./script3_extra.json");
const RU = require("./content3.json");

const OUT = process.argv[2] || path.join(__dirname, "speech_en.docx");
const FONT = "Times New Roman";
const PAGE_W = 11906, PAGE_H = 16838, MARGIN = 1134; // A4, 2 cm margins
const CONTENT_W = PAGE_W - 2 * MARGIN; // 9638
const SHADE = "E7E6E6"; // speaker 2 rows — visible in grayscale

const run = (text, o = {}) => new TextRun({ text, font: FONT, size: o.size || 28, bold: o.bold, italics: o.italics, color: o.color, allCaps: o.caps });
const para = (children, o = {}) => new Paragraph({ children: Array.isArray(children) ? children : [children], alignment: o.align, spacing: { before: o.before || 0, after: o.after ?? 120, line: o.line || 312 }, keepNext: o.keepNext, keepLines: o.keepLines, heading: o.heading, border: o.border, indent: o.indent });
const none = { style: BorderStyle.NONE, size: 0, color: "FFFFFF" };
const thin = { style: BorderStyle.SINGLE, size: 4, color: "000000" };
const cell = (children, w, o = {}) => new TableCell({
  children, width: { size: w, type: WidthType.DXA }, verticalAlign: o.v || VerticalAlign.TOP,
  shading: o.shade ? { type: ShadingType.CLEAR, color: "auto", fill: o.shade } : undefined,
  margins: { top: o.pad ?? 100, bottom: o.pad ?? 100, left: 140, right: 140 },
  borders: o.borders || { top: thin, bottom: thin, left: thin, right: thin },
  columnSpan: o.span,
});

const speakerOf = (sp) => (String(sp).startsWith("Nikita") ? { name: "NIKITA", shade: undefined } : { name: "ILYA", shade: SHADE });
const ruTitles = ["Титульный слайд"].concat(RU.slides.map((x) => x.title || "Спасибо за внимание!"));
SLIDES.forEach((s) => { s.ru = ruTitles[s.n - 1]; });
const fmtTime = (s) => `${Math.floor(s / 60)}:${String(s % 60).padStart(2, "0")}`;

const children = [];

// ---------- title block ----------
children.push(para(run("PRESENTATION SCRIPT", { size: 24, bold: true, color: "595959" }), { align: AlignmentType.CENTER, after: 60 }));
children.push(para(run("Cosmic Radiation Exposure of Aircraft Crew Members and Their Legal Protection in the Russian Federation", { size: 34, bold: true }), { align: AlignmentType.CENTER, after: 120, line: 300 }));
children.push(para(run("Космическое облучение членов экипажей воздушных судов и их правовая защита в Российской Федерации", { size: 24, italics: true }), { align: AlignmentType.CENTER, after: 240 }));

const infoRows = [
  ["Speaker 1", "Nikita Shutov (Шутов Н. Р.) — “NIKITA” blocks, white background"],
  ["Speaker 2", "Ilya Roldugin (Ролдугин И. Д.) — “ILYA” blocks, grey background"],
  ["Supervisor", "Professor Natalia S. Boyko, Department of Flight Operations and Flight Safety"],
  ["Language", "Simple English (level B1): short sentences, common words"],
  ["Slides", "The slides are in Russian; this English text follows them slide by slide"],
  ["Timing", `About ${Math.round(SLIDES.reduce((a, s) => a + s.sec, 0) / 60)} minutes at a calm pace (≈ ${WPM} words per minute). Rehearse aloud with a timer.`],
];
children.push(new Table({
  width: { size: CONTENT_W, type: WidthType.DXA }, columnWidths: [2200, CONTENT_W - 2200], layout: TableLayoutType.FIXED,
  rows: infoRows.map(([a, b], i) => new TableRow({ children: [
    cell([para(run(a, { size: 22, bold: true }), { after: 0, line: 260 })], 2200, { shade: i === 1 ? SHADE : "F2F2F2", pad: 60 }),
    cell([para(run(b, { size: 22 }), { after: 0, line: 260 })], CONTENT_W - 2200, { shade: i === 1 ? SHADE : undefined, pad: 60 }),
  ] })),
}));

// ---------- running order ----------
children.push(para(run("Running order", { size: 30, bold: true }), { before: 300, after: 120, keepNext: true }));
const ro = [900, 6338, 1400, 1000];
children.push(new Table({
  width: { size: CONTENT_W, type: WidthType.DXA }, columnWidths: ro, layout: TableLayoutType.FIXED,
  rows: [
    new TableRow({ tableHeader: true, children: ["Slide", "Title (exact slide title in Russian)", "Speaker", "Time"].map((h, i) =>
      cell([para(run(h, { size: 22, bold: true }), { after: 0, align: i === 1 ? AlignmentType.LEFT : AlignmentType.CENTER })], ro[i], { shade: "D9D9D9" })) }),
    ...SLIDES.map((s) => {
      const sp = speakerOf(s.parts[0][0]);
      return new TableRow({ children: [
        cell([para(run(String(s.n), { size: 22, bold: true }), { after: 0, align: AlignmentType.CENTER, line: 260 })], ro[0], { shade: sp.shade, pad: 50 }),
        cell([para([run(s.title, { size: 22 }), run(`  (${s.ru})`, { size: 20, italics: true, color: "404040" })], { after: 0, line: 260 })], ro[1], { shade: sp.shade, pad: 50 }),
        cell([para(run(sp.name, { size: 22, bold: true }), { after: 0, align: AlignmentType.CENTER, line: 260 })], ro[2], { shade: sp.shade, pad: 50 }),
        cell([para(run(fmtTime(s.sec), { size: 22 }), { after: 0, align: AlignmentType.CENTER, line: 260 })], ro[3], { shade: sp.shade, pad: 50 }),
      ] });
    }),
  ],
}));

// ---------- script ----------
children.push(new Paragraph({ children: [new PageBreak()] }));
children.push(para(run("Script", { size: 30, bold: true }), { before: 0, after: 60, keepNext: true }));
children.push(para(run("Bold arrows (►) mark the hand-over to the other speaker. “[Next slide]” marks the moment to switch the slide.", { size: 22, italics: true, color: "404040" }), { after: 160 }));
const sw = [1650, CONTENT_W - 1650];
SLIDES.forEach((s, idx) => {
  s.parts.forEach(([spk, text]) => {
    const sp = speakerOf(spk);
    const next = SLIDES[idx + 1];
    const handover = next && next.parts[0][0].key !== spk.key;
    const body = [para([run(text)], { after: 80, line: 360 })];
    const cue = next ? (handover ? `► ${speakerOf(next.parts[0][0]).name} continues  ·  [Next slide]` : "[Next slide]") : "End of the talk — questions";
    body.push(para(run(cue, { size: 22, italics: !handover, bold: handover, color: handover ? "000000" : "595959" }), { after: 0, align: AlignmentType.RIGHT }));
    children.push(new Table({
      width: { size: CONTENT_W, type: WidthType.DXA }, columnWidths: sw, layout: TableLayoutType.FIXED,
      rows: [new TableRow({ cantSplit: true, children: [
        cell([
          para(run(`SLIDE ${s.n}`, { size: 20, bold: true, color: "404040" }), { after: 40 }),
          para(run(sp.name, { size: 30, bold: true }), { after: 40 }),
          para(run(fmtTime(s.sec), { size: 20, color: "404040" }), { after: 0 }),
        ], sw[0], { shade: sp.shade }),
        cell([
          para([run(s.title, { size: 24, bold: true }), run(`  (${s.ru})`, { size: 20, italics: true, color: "404040" })], { after: 80, keepNext: true }),
          ...body,
        ], sw[1], { shade: sp.shade }),
      ] })],
    }));
    children.push(para(run(""), { after: 60, line: 240 }));
  });
});

// ---------- pronunciation ----------
children.push(para(run("Terms and how to read them", { size: 30, bold: true }), { before: 300, after: 120, keepNext: true }));
const gw = [3000, CONTENT_W - 3000];
children.push(new Table({
  width: { size: CONTENT_W, type: WidthType.DXA }, columnWidths: gw, layout: TableLayoutType.FIXED,
  rows: [
    new TableRow({ tableHeader: true, children: ["Written", "Read as / meaning"].map((h, i) => cell([para(run(h, { size: 22, bold: true }), { after: 0 })], gw[i], { shade: "D9D9D9" })) }),
    ...EXTRA.terms.map(([a, b]) => new TableRow({ children: [
      cell([para(run(a, { size: 22, bold: true }), { after: 0 })], gw[0]),
      cell([para(run(b, { size: 22 }), { after: 0 })], gw[1]),
    ] })),
  ],
}));

// ---------- Q&A ----------
children.push(para(run("Possible questions from the jury", { size: 30, bold: true }), { before: 360, after: 120, keepNext: true }));
EXTRA.qa.forEach((q, i) => {
  const sp = speakerOf(q.who);
  children.push(new Table({
    width: { size: CONTENT_W, type: WidthType.DXA }, columnWidths: sw, layout: TableLayoutType.FIXED,
    rows: [new TableRow({ cantSplit: true, children: [
      cell([para(run(`Q${i + 1}`, { size: 20, bold: true, color: "404040" }), { after: 40 }), para(run(sp.name, { size: 26, bold: true }), { after: 0 })], sw[0], { shade: sp.shade }),
      cell([
        para([run("Q: ", { size: 24, bold: true }), run(q.q, { size: 24, bold: true })], { after: 80, keepNext: true }),
        para([run("A: ", { size: 24, bold: true }), run(q.a, { size: 24 })], { after: 0, line: 336 }),
      ], sw[1], { shade: sp.shade }),
    ] })],
  }));
  children.push(para(run(""), { after: 60, line: 240 }));
});

const doc = new Document({
  creator: "Shutov N. R., Roldugin I. D.",
  title: "Presentation script — Cosmic Radiation Exposure of Aircraft Crew Members",
  styles: { default: { document: { run: { font: FONT, size: 28 } } } },
  sections: [{
    properties: { page: { size: { width: PAGE_W, height: PAGE_H }, margin: { top: MARGIN, bottom: MARGIN, left: MARGIN, right: MARGIN } } },
    headers: { default: new Header({ children: [para(run("Presentation script · Shutov N. R., Roldugin I. D. · Ulyanovsk Civil Aviation Institute, 2026", { size: 18, color: "595959" }), { align: AlignmentType.RIGHT, after: 0 })] }) },
    footers: { default: new Footer({ children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ children: [PageNumber.CURRENT], font: FONT, size: 20 })] })] }) },
    children,
  }],
});

Packer.toBuffer(doc).then((buf) => { fs.writeFileSync(OUT, buf); console.log("written", OUT); });
