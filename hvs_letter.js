// HvS Attorneys letter renderer. Usage: node hvs_letter.js spec.json
// Needs: docx (npm), logo at ~/letterhead/logo.png (written from the skill).
const fs = require("fs"), os = require("os"), path = require("path");
const { Document, Packer, Paragraph, TextRun, ImageRun, Table, TableRow, TableCell, AlignmentType,
  BorderStyle, WidthType, ShadingType, Footer, PageBreak, TabStopType } = require("docx");

const spec = JSON.parse(fs.readFileSync(process.argv[2], "utf8"));
const LOGO = fs.readFileSync(spec.logo || path.join(os.homedir(), "letterhead", "logo.png"));
const GOLD = "AE8A3D", FONT = "Arial", SZ = 20, CW = 9306;
const TEL = "010 141 2084";

// "**bold**" and "*italic*" inside strings become runs
function runs(s, base = {}) {
  const out = [];
  String(s).split(/(\*\*[^*]+\*\*|\*[^*]+\*)/).forEach(part => {
    if (!part) return;
    if (part.startsWith("**")) out.push(R(part.slice(2, -2), { ...base, b: true }));
    else if (part.startsWith("*")) out.push(R(part.slice(1, -1), { ...base, i: true }));
    else out.push(R(part, base));
  });
  return out;
}
function R(t, o = {}) {
  return new TextRun({ text: t, bold: o.b, italics: o.i, font: FONT, size: o.size || SZ, color: o.color });
}
const P = (children, o = {}) => new Paragraph({
  alignment: o.align || AlignmentType.LEFT,
  spacing: o.noLine ? { before: 0, after: o.after ?? 0 }
                    : { before: o.before || 0, after: o.after ?? 120, line: o.line || 276 },
  indent: o.indent, tabStops: o.tabs, border: o.border, keepNext: o.keepNext, children });
const goldRule = keepNext => P([R("")], { after: 200, keepNext,
  border: { bottom: { color: GOLD, space: 1, style: BorderStyle.SINGLE, size: 12 } } });
const logo = (w, h) => new ImageRun({ type: "png", data: LOGO, transformation: { width: w, height: h } });
const none = { style: BorderStyle.NONE, size: 0, color: "FFFFFF" };
const noBorders = { top: none, bottom: none, left: none, right: none, insideHorizontal: none, insideVertical: none };

// letterhead: logo left, contact line right
const head = new Table({
  width: { size: CW, type: WidthType.DXA }, columnWidths: [2700, CW - 2700], layout: "fixed", borders: noBorders,
  rows: [new TableRow({ children: [
    new TableCell({ width: { size: 2700, type: WidthType.DXA }, borders: noBorders,
      children: [P([logo(178, 124)], { noLine: true })] }),
    new TableCell({ width: { size: CW - 2700, type: WidthType.DXA }, borders: noBorders, verticalAlign: "bottom",
      children: [P([R(`35 Fricker Rd, Illovo, Sandton | ${TEL}`, { size: 16 })], { align: AlignmentType.RIGHT, after: 0 }),
        P([R("enquiries@hvslegal.co.za | www.hvslegal.co.za", { size: 16 })], { align: AlignmentType.RIGHT, after: 60 })] }),
  ] })],
});

const body = [head, goldRule()];
if (spec.our_ref !== undefined && spec.our_ref !== null)
  body.push(P([R("OUR REF: ", { b: true }), R(spec.our_ref || ""), R("\tYOUR REF: ", { b: true }), R(spec.your_ref || "")],
    { tabs: [{ type: TabStopType.LEFT, position: 5000 }] }));
body.push(P([R("Date: ", { b: true }), R(spec.date)], { after: 200 }));
if (spec.to && spec.to.length) {
  body.push(P([R("TO:", { b: true })], { after: 40 }));
  spec.to.forEach((l, i) => body.push(P(runs(l), { after: i === spec.to.length - 1 ? 200 : 0 })));
}
body.push(P(runs(spec.salutation || "Dear Sir/Madam,"), { after: 160 }));
body.push(goldRule(true));
body.push(P([R(spec.re, { b: true })], { align: AlignmentType.CENTER, keepNext: true, after: 0 }));
body.push(goldRule());

function table(t) {
  const n = t.header ? t.header.length : t.rows[0].length;
  const widths = t.widths || Array(n).fill(Math.floor(CW / n));
  const b = { top: { style: BorderStyle.SINGLE, size: 4, color: GOLD }, bottom: { style: BorderStyle.SINGLE, size: 4, color: GOLD },
    left: none, right: none, insideVertical: none, insideHorizontal: { style: BorderStyle.SINGLE, size: 2, color: "DDDDDD" } };
  const row = (cells, hdr) => new TableRow({ tableHeader: hdr, children: cells.map((c, i) => new TableCell({
    width: { size: widths[i], type: WidthType.DXA },
    shading: hdr ? { type: ShadingType.CLEAR, fill: "F2EEE6", color: "auto" } : undefined,
    margins: { top: 70, bottom: 70, left: 100, right: 100 },
    children: [P(runs(String(c), hdr ? { b: true } : {}), { after: 0 })] })) });
  return new Table({ width: { size: CW, type: WidthType.DXA }, columnWidths: widths, layout: "fixed", borders: b,
    rows: [...(t.header ? [row(t.header, true)] : []), ...t.rows.map(r => row(r, false))] });
}

for (const blk of spec.body) {
  if (blk.h) body.push(P(runs(blk.h, { b: true }), { before: 240, after: 140, keepNext: true }));
  else if (blk.n) {
    const lvl = String(blk.n).split(".").filter(Boolean).length;
    const ind = lvl === 1 ? { left: 700, hanging: 700 } : lvl === 2 ? { left: 1560, hanging: 860 } : { left: 2560, hanging: 1000 };
    const num = lvl === 1 && !String(blk.n).endsWith(".") ? blk.n + "." : blk.n;
    body.push(P([R(num + "\t"), ...runs(blk.text)], { align: AlignmentType.JUSTIFIED, after: 140, line: 300,
      indent: ind, tabs: [{ type: TabStopType.LEFT, position: ind.left }] }));
  }
  else if (blk.p !== undefined) body.push(P(runs(blk.p), { align: AlignmentType.JUSTIFIED, after: 140, line: 300 }));
  else if (blk.table) { body.push(table(blk.table)); body.push(P([R("")], { after: 120 })); }
}

if (spec.type === "engagement") body.push(new Paragraph({ children: [new PageBreak()] }));
body.push(P([R(spec.signoff || "Yours faithfully,")], { before: 200, keepNext: true }));
body.push(P([R("")], { after: 400, keepNext: true }));
const sigRef = spec.our_ref ? ` | Ref: ${spec.our_ref}` : "";
[["HvS ATTORNEYS", true], ["Per: H VAN SCHALKWYK", false], ["Director | B.Soc.Sci (cum laude) LLB", false],
 [`Tel: ${TEL} | director@hvslegal.co.za${sigRef}`, false]]
  .forEach(([t, b], i, a) => body.push(P([R(t, { b })], { after: 0, keepNext: i < a.length - 1 })));

if (spec.type === "engagement" && spec.trust !== false) {
  body.push(P([R("")], { after: 200 }));
  body.push(P([R("Trust account details", { b: true })], { keepNext: true }));
  body.push(table({ widths: [3000, CW - 3000], rows: [
    ["Account name", "Commercial Attorneys Trust Account"], ["Bank", spec.bank || "FNB"],
    ["Account number", "63145380985"], ["Payment reference", spec.our_ref || ""]] }));
}
(spec.after_signature || []).forEach(blk => {
  if (blk.h) body.push(P(runs(blk.h, { b: true }), { before: 240, after: 140, keepNext: true }));
  else if (blk.table) body.push(table(blk.table));
  else body.push(P(runs(blk.p || ""), { after: 140 }));
});

const footer = new Footer({ children: [
  P([R("")], { after: 80, border: { top: { color: GOLD, space: 1, style: BorderStyle.SINGLE, size: 12 } } }),
  P([logo(77, 54)], { align: AlignmentType.CENTER, noLine: true }),
  P([R("Director: Hyran van Schalkwyk", { i: true, color: "777777", size: 16 })], { align: AlignmentType.CENTER, after: 0 }),
] });

const doc = new Document({ sections: [{
  properties: { page: { size: { width: 11906, height: 16838 }, margin: { top: 900, bottom: 760, left: 1300, right: 1300, footer: 300 } } },
  footers: { default: footer }, children: body }] });
Packer.toBuffer(doc).then(b => { fs.writeFileSync(spec.out || "out.docx", b); console.log("wrote", spec.out || "out.docx"); });
