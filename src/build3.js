// Доклад по воздушному праву — оформление по правилам ГОСТ 7.32-2017 для слайдов:
// Times New Roman, заголовки по центру без точки, перечисления через тире,
// «Рисунок N – …» под рисунком, «Таблица N – …» над таблицей, номер слайда справа внизу.
const path = require("path");
const pptxgen = require("pptxgenjs");
const fs = require("fs");
const SKILL = [
  "/root/.claude/skills/synced/befae669-fc54-4426-a37b-2f5acf0e68ba_3ce5ac46-9f43-4ba0-887b-072c4916456d/pptx",
  "/mnt/skills/public/pptx",
].find((d) => fs.existsSync(path.join(d, "scripts/apply_theme.js")));
const { applyTheme } = require(path.join(SKILL, "scripts/apply_theme.js"));
const K = require("./content3.json");
const { SLIDES: NOTES } = require("./notes3_en.js");

const IMG = (f) => path.join(__dirname, "img", f);
const OUT = process.argv[2] || path.join(__dirname, "deck3.pptx");
const STYLE = process.argv[3] || "gost";
const NICE = STYLE === "nice";

const THEME = {
  name: "Office",
  headFontFace: NICE ? "Cambria" : "Arial",
  bodyFontFace: NICE ? "Calibri" : "Arial",
  colors: {
    dk1: NICE ? "1A1A1A" : "000000", lt1: "FFFFFF", dk2: NICE ? "1F3864" : "000000", lt2: NICE ? "EAF0F8" : "FFFFFF",
    accent1: NICE ? "1F3864" : "000000", accent2: NICE ? "8EA9DB" : "7F7F7F", accent3: "C00000",
    accent4: "BFBFBF", accent5: "595959", accent6: "D9D9D9",
    hlink: "0563C1", folHlink: "954F72",
  },
};
const HEX = THEME.colors;

const pres = new pptxgen();
pres.layout = "LAYOUT_WIDE";
pres.theme = { headFontFace: THEME.headFontFace, bodyFontFace: THEME.bodyFontFace };
pres.title = K.titlePlain;
pres.author = "Шутов Н. Р., Ролдугин И. Д.";
pres.company = "УИ ГА";
pres.subject = "Воздушное право";

const C = pres.SchemeColor;
const BLACK = C.text1, WHITE = C.background1, GREY = C.accent2, RED = C.accent3, NAVY = C.text2, PALE = C.background2;
const TFONT = NICE ? "Calibri" : "Arial";
const SW = 13.333, ML = 0.6, CW = SW - 2 * ML;

// неразрывные пробелы и дефисы
const NB = " ";
const nb = (s) => s
  .replace(/(\d)\s+(?=(?:мкЗв|мЗв|мкР|Зв|ч|часов|часа|час|раз|раза|фут|%|г\.|года|году|месяц|месяца|месяцев)(?![А-Яа-яЁё]))/g, "$1" + NB)
  .replace(/(^|[\s(«–])([вскиоуаВСКИОУА])\s+/g, "$1$2" + NB)
  .replace(/(№|ст\.|п\.|пп\.|ч\.|т\.|гл\.|§)\s+/g, "$1" + NB)
  .replace(/\s+–/g, NB + "–")
  .replace(/(\d)\s+(\d{3})(?!\d)/g, "$1" + NB + "$2")
  .replace(/CARI-7/g, "CARI‑7")
  .replace(/(\d)-(\d{2})(?!\d)/g, "$1‑$2")
  .replace(/НРБ-99/g, "НРБ‑99")
  .replace(/ОСПОРБ-99/g, "ОСПОРБ‑99")
  .replace(/(\d)-ФЗ/g, "$1‑ФЗ")
  .replace(/(ТК|ВК|кодекс|кодекса) РФ/g, "$1" + NB + "РФ")
  .replace(/≈\s+/g, "≈" + NB)
  .replace(/\s+\[(\d)/g, NB + "[$1")
  .replace(/(\d)–(\d)/g, "$1–\u2060$2");
const fix = (t) => (typeof t === "string" ? nb(t) : t.map((r) => ({ ...r, text: nb(r.text) })));

// ---------- макеты (как стандартные макеты PowerPoint) ----------
const TITLE_PH = () => ({ name: "title", type: "title", x: 1.45, y: 0.3, w: SW - 2.9, h: 1.0, fontSize: NICE ? 28 : 32, bold: true, color: NICE ? NAVY : BLACK, align: NICE ? "left" : "center", valign: "middle", margin: 0.05, lang: "ru-RU" });
const NUM = () => ({ x: SW - 1.3, y: 6.95, w: 0.8, h: 0.35, fontSize: 14, color: BLACK, align: "right" });
pres.defineSlideMaster({ title: "Титульный слайд", background: { color: WHITE } });
const LOGO = () => ({ image: { path: IMG("logo.png"), x: 0.4, y: 0.35, w: 0.9, h: 0.9, altText: "Эмблема УИ ГА" } });
const RULE = () => ({ line: { x: 1.45, y: 1.22, w: SW - 2.05, h: 0, line: { color: NAVY, width: 1.5 } } });
const extra = () => (NICE ? [LOGO(), RULE()] : [LOGO()]);
pres.defineSlideMaster({
  title: "Заголовок и объект", background: { color: WHITE },
  objects: [
    ...extra(),
    { placeholder: { options: TITLE_PH(), text: "" } },
    { placeholder: { options: { name: "body", type: "body", x: ML, y: 1.5, w: CW, h: 5.3, fontSize: 24, color: BLACK, align: "left", valign: "top", margin: 0.05, lang: "ru-RU" }, text: "" } },
  ],
  slideNumber: NUM(),
});
pres.defineSlideMaster({
  title: "Только заголовок", background: { color: WHITE },
  objects: [...extra(), { placeholder: { options: TITLE_PH(), text: "" } }],
  slideNumber: NUM(),
});

// ---------- помощники ----------
const T = (s, text, o) => s.addText(fix(text), { isTextBox: true, margin: 0.05, color: BLACK, fontSize: 20, valign: "top", lang: "ru-RU", ...o });
const title = (s, t) => s.addText(nb(t), { placeholder: "title" });
// перечисление через тире (ГОСТ 7.32): элементы — строки или {text, sub:[...]} ; можно {lead:'…', text:'…'} для полужирного начала
const dashList = (items, fs = 20) => {
  const runs = [];
  items.forEach((it, i) => {
    const last = i === items.length - 1;
    const para = { bullet: { characterCode: "2013", indent: 24 }, paraSpaceAfter: 10, align: "left" };
    if (typeof it === "string") runs.push({ text: nb(it), options: { ...para, breakLine: !last } });
    else {
      if (it.lead) runs.push({ text: nb(it.lead) + " ", options: { ...para, bold: true } });
      runs.push({ text: nb(it.text), options: { ...para, breakLine: !last } });
    }
  });
  return runs;
};
const bodyList = (s, items, fs) => s.addText(dashList(items, fs), { placeholder: "body", fontSize: fs || 24 });
const listBox = (s, items, o) => s.addText(dashList(items), { isTextBox: true, margin: 0.05, color: BLACK, fontSize: 20, valign: "top", lang: "ru-RU", ...o });
const caption = (s, text, x, y, w, h = 0.4) => T(s, text, { x, y, w, h, fontSize: 16, align: "center", objectName: "Подпись" });
const picture = (s, file, x, y, w, h, alt) => s.addImage({ path: IMG(file), x, y, w, h, altText: alt });
const rect = (s, x, y, w, h, dash) => s.addShape(NICE ? pres.ShapeType.roundRect : pres.ShapeType.rect, { x, y, w, h, ...(NICE ? { rectRadius: 0.08 } : {}), fill: { color: NICE && !dash ? PALE : WHITE }, line: { color: NICE ? NAVY : BLACK, width: 1, ...(dash ? { dashType: "dash" } : {}) } });
const arrow = (s, x1, y1, x2, y2) => {
  const o = { line: { color: NICE ? NAVY : BLACK, width: 1.25, endArrowType: "triangle" } };
  if (x2 >= x1) s.addShape(pres.ShapeType.line, { x: x1, y: y1, w: x2 - x1, h: y2 - y1, ...o });
  else s.addShape(pres.ShapeType.line, { x: x2, y: y1, w: x1 - x2, h: y2 - y1, flipH: true, ...o });
};
const notes = (s, n) => {
  const e = NOTES.find((x) => x.n === n);
  if (e) s.addNotes(e.parts.map(([sp, t]) => `${sp.toUpperCase()}:\n${t}`).join("\n\n"));
};
const table = (s, rows, o) => {
  const border = { type: "solid", pt: NICE ? 0.75 : 1, color: NICE ? "8EA9DB" : "000000" };
  const data = rows.map((r, ri) => r.map((c) => {
    const cell = typeof c === "string" ? { text: c } : c;
    const text = typeof cell.text === "string" ? nb(cell.text) : cell.text;
    const look = NICE ? (ri === 0 ? { fill: { color: NAVY }, color: WHITE } : (ri % 2 === 0 ? { fill: { color: PALE } } : {})) : {};
    return { text, options: { bold: ri === 0, align: cell.align || (ri === 0 ? "center" : "left"), valign: "middle", ...look, ...(cell.options || {}) } };
  }));
  s.addTable(data, { fontFace: TFONT, fontSize: 16, color: BLACK, border, margin: [0.04, 0.08, 0.04, 0.08], autoPage: false, lang: "ru-RU", ...o });
};

let n = 0;
const slide = (layout) => { n += 1; return pres.addSlide({ masterName: layout }); };

// ================= 1. Титульный слайд =================
{
  const s = slide("Титульный слайд");
  s.addImage({ path: IMG("logo.png"), x: 0.45, y: 0.3, w: 1.0, h: 1.0, altText: "Эмблема УИ ГА" });
  T(s, K.header.map((t, i) => ({ text: t, options: { breakLine: i < K.header.length - 1 } })), { x: 1.55, y: 0.3, w: 10.6, h: 1.55, fontSize: 13, align: "center", lineSpacingMultiple: 1.0 });
  T(s, K.department, { x: 1.6, y: 1.95, w: 10.13, h: 0.4, fontSize: 16, align: "center" });
  T(s, "ДОКЛАД", { x: 1.6, y: 2.45, w: 10.13, h: 0.45, fontSize: 20, bold: true, align: "center" });
  T(s, K.titleQuoted, { x: 0.8, y: 2.9, w: 11.73, h: 1.2, fontSize: 24, bold: true, align: "center", valign: "middle", ...(NICE ? { color: NAVY, fontFace: "Cambria" } : {}) });
  T(s, `по дисциплине «${K.discipline}»`, { x: 1.6, y: 4.05, w: 10.13, h: 0.4, fontSize: 16, align: "center" });
  const sig = [];
  K.signature.forEach((l, i) => sig.push({ text: l.t, options: { bold: !!l.b, breakLine: i < K.signature.length - 1 } }));
  T(s, sig, { x: 7.3, y: 4.6, w: 5.45, h: 1.95, fontSize: 16, align: "left" });
  T(s, "Ульяновск 2026", { x: 4.67, y: 6.85, w: 4.0, h: 0.4, fontSize: 16, align: "center" });
  notes(s, n);
}

// ================= контентные слайды из content3.json =================
for (const sl of K.slides) {
  if (sl.type === "bullets") {
    const s = slide("Заголовок и объект");
    title(s, sl.title);
    bodyList(s, sl.items, sl.fontSize);
    notes(s, n);
  } else if (sl.type === "goals") {
    const s = slide("Только заголовок");
    title(s, sl.title);
    const rows = [];
    sl.rows.forEach(([h, t], i) => {
      rows.push({ text: nb(h) + " ", options: { bold: true } });
      rows.push({ text: nb(t), options: { breakLine: i < sl.rows.length - 1 } });
    });
    T(s, rows.map((r, i) => ({ ...r, options: { ...r.options, paraSpaceAfter: 10, align: "left" } })), { x: ML, y: 1.5, w: CW, h: 2.4, fontSize: 24 });
    T(s, sl.tasksTitle, { x: ML, y: 3.95, w: CW, h: 0.45, fontSize: 24, bold: true });
    s.addText(sl.tasks.map((t, i) => ({ text: nb(t), options: { bullet: { type: "number", style: "arabicParenR" }, paraSpaceAfter: 4, breakLine: i < sl.tasks.length - 1 } })),
      { isTextBox: true, x: ML, y: 4.45, w: CW, h: 2.45, fontSize: 24, color: BLACK, valign: "top", margin: 0.05, lang: "ru-RU" });
    notes(s, n);
  } else if (sl.type === "table-measure") {
    const s = slide("Только заголовок");
    title(s, sl.title);
    const tw = 7.3;
    T(s, sl.tableTitle, { x: ML, y: 1.4, w: tw, h: 0.45, fontSize: 20 });
    const rows = [sl.head.map((h) => ({ text: h, align: "center" }))].concat(sl.rows.map((r) => r.map((c, ci) => ({ text: c, align: ci === 0 ? "left" : "center" }))));
    table(s, rows, { x: ML, y: 1.9, w: tw, colW: [2.4, 1.55, 1.55, 1.8], rowH: 0.55, fontSize: 18 });
    listBox(s, sl.items, { x: ML + tw + 0.35, y: 1.9, w: CW - tw - 0.35, h: 4.9, fontSize: 22 });
    notes(s, n);
  } else if (sl.type === "photos") {
    const s = slide("Только заголовок");
    title(s, sl.title);
    const pw = 2.85, ph = 4.04, g = (CW - 4 * pw) / 3, y = 1.35;
    sl.photos.forEach(([f, lab, alt], i) => {
      const x = ML + i * (pw + g);
      picture(s, f, x, y, pw, ph, alt);
      T(s, lab, { x: x - 0.1, y: y + ph + 0.05, w: pw + 0.2, h: 0.35, fontSize: 14, align: "center" });
    });
    caption(s, sl.caption, ML, 5.95, CW, 0.8);
    notes(s, n);
  } else if (sl.type === "scheme") {
    const s = slide("Только заголовок");
    title(s, sl.title);
    const box = (x, y, w, h, head, body, dash) => {
      rect(s, x, y, w, h, dash);
      T(s, [{ text: head, options: { bold: true, breakLine: true } }, { text: body }], { x: x + 0.1, y: y + 0.04, w: w - 0.2, h: h - 0.08, fontSize: 16, align: "center", valign: "middle" });
    };
    const top = { x: 1.9, y: 1.35, w: 9.53, h: 0.85 };
    box(top.x, top.y, top.w, top.h, sl.top[0], sl.top[1]);
    const mw = 3.75, mh = 1.2, my = 2.65, mg = (CW - 3 * mw) / 2;
    const bot = { x: 4.42, y: 4.3, w: 4.5, h: 0.75 };
    sl.mid.forEach(([h, t], i) => {
      const x = ML + i * (mw + mg), cx = x + mw / 2;
      box(x, my, mw, mh, h, t);
      arrow(s, cx, top.y + top.h, cx, my);
      const bx = bot.x + bot.w * (0.2 + 0.3 * i);
      arrow(s, cx, my + mh, bx, bot.y);
    });
    box(bot.x, bot.y, bot.w, bot.h, sl.bottom[0], sl.bottom[1]);
    const dw = 5.85, dy = 5.35, dh = 0.85;
    box(ML, dy, dw, dh, sl.lost[0], sl.lost[1], true);
    box(SW - ML - dw, dy, dw, dh, sl.gap[0], sl.gap[1], true);
    caption(s, sl.caption, ML, 6.3, CW, 0.65);
    notes(s, n);
  } else if (sl.type === "chart") {
    const s = slide("Только заголовок");
    title(s, sl.title);
    const lw = 5.6;
    listBox(s, sl.items, { x: ML, y: 1.4, w: lw, h: 5.4, fontSize: 20 });
    const x = ML + lw + 0.4, w = SW - ML - x;
    s.addChart(pres.ChartType.bar, [{ name: "Годовая доза, мЗв", labels: sl.chart.labels, values: sl.chart.values }], {
      x, y: 1.5, w, h: 4.3, barDir: "col",
      chartColors: [HEX.accent4, HEX.accent5, HEX.dk1], barGapWidthPct: 60,
      showValue: false, catAxisLabelFontSize: 14, catAxisLabelColor: "000000", catAxisLabelFontFace: "+mn-lt",
      valAxisLabelFontSize: 14, valAxisLabelColor: "000000", valAxisLabelFontFace: "+mn-lt", valAxisMinVal: 0, valAxisMaxVal: 6, valAxisMajorUnit: 1,
      valAxisTitle: "мЗв в год", showValAxisTitle: true, valAxisTitleFontSize: 14, valAxisTitleColor: "000000", valAxisTitleFontFace: "+mn-lt",
      valGridLine: { color: "BFBFBF", size: 0.5 }, catGridLine: { style: "none" }, showLegend: false, showTitle: false,
      altText: sl.chart.alt,
    });
    caption(s, sl.caption, x, 5.9, w, 0.75);
    notes(s, n);
  } else if (sl.type === "compare") {
    const s = slide("Только заголовок");
    title(s, sl.title);
    T(s, sl.tableTitle, { x: ML, y: 1.4, w: CW, h: 0.45, fontSize: 20 });
    const rows = [sl.head.map((h) => ({ text: h, align: "center" }))].concat(sl.rows);
    table(s, rows, { x: ML, y: 1.9, w: CW, colW: sl.colW, fontSize: sl.fontSize || 18, rowH: sl.rowH });
    if (sl.note) T(s, sl.note, { x: ML, y: sl.noteY || 5.6, w: CW, h: 0.9, fontSize: 20, bold: true });
    notes(s, n);
  } else if (sl.type === "figures") {
    const s = slide("Только заголовок");
    title(s, sl.title);
    if (sl.stack) {
      let y = 1.45; const w = sl.stackW || 9.5, x = ML + (CW - w) / 2;
      sl.figs.forEach((f) => { const h = w / f.ar; picture(s, f.file, x, y, w, h, f.alt); y += h + 0.15; });
      caption(s, sl.caption, ML, y, CW, 0.7);
      notes(s, n);
      continue;
    }
    const top = 1.45, maxH = sl.maxH || 4.55, gap = 0.4;
    const ws = sl.figs.map((f) => Math.min(maxH * f.ar, (CW - gap * (sl.figs.length - 1)) * (f.share || 1 / sl.figs.length)));
    const total = ws.reduce((a, w) => a + w, 0) + gap * (sl.figs.length - 1);
    let x = ML + (CW - total) / 2;
    sl.figs.forEach((f, i) => {
      const w = ws[i], h = w / f.ar;
      picture(s, f.file, x, top, w, h, f.alt);
      if (f.caption) caption(s, f.caption, x - 0.1, top + h + 0.08, w + 0.2, 0.75);
      x += w + gap;
    });
    if (sl.caption) caption(s, sl.caption, ML, sl.captionY || 6.2, CW, 0.7);
    notes(s, n);
  } else if (sl.type === "refs") {
    const s = slide("Только заголовок");
    title(s, sl.title);
    s.addText(sl.refs.map((r, i) => ({ text: nb(r), options: { bullet: { type: "number", numberStartAt: sl.start || 1 }, paraSpaceAfter: 4, align: "justify", breakLine: i < sl.refs.length - 1 } })),
      { isTextBox: true, x: ML, y: 1.4, w: CW, h: 5.45, fontSize: sl.fontSize || 14, color: BLACK, valign: "top", margin: 0.05, lang: "ru-RU" });
    notes(s, n);
  } else if (sl.type === "thanks") {
    const s = slide("Титульный слайд");
    s.addImage({ path: IMG("logo.png"), x: SW / 2 - 0.6, y: 1.2, w: 1.2, h: 1.2, altText: "Эмблема УИ ГА" });
    T(s, "Спасибо за внимание!", { x: ML, y: 2.6, w: CW, h: 1.0, fontSize: 36, bold: true, align: "center", valign: "middle", ...(NICE ? { color: NAVY, fontFace: "Cambria" } : {}) });
    T(s, K.titleQuoted, { x: 1.2, y: 3.7, w: 10.93, h: 1.0, fontSize: 20, align: "center", valign: "middle" });
    notes(s, n);
  }
}

(async () => {
  await pres.writeFile({ fileName: OUT });
  await applyTheme(OUT, THEME);
  console.log("written", OUT, n, "slides");
})();
