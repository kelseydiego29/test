const fs = require('fs');
const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell, WidthType,
  AlignmentType, BorderStyle, PageBreak, Footer, VerticalAlign,
} = require('docx');

/* ------------------------------------------------------------------
   One typeface, one colour. Emphasis is carried by weight, size and
   spacing only — no second font, no accent colour, no shading.
   ------------------------------------------------------------------ */
const FONT = 'Georgia';
const BLACK = '000000';

/* A4 with 19mm margins */
const PAGE_W = 11906, PAGE_H = 16838, MARGIN = 1080;
const W = PAGE_W - MARGIN * 2; // 9746

const NONE = { style: BorderStyle.NONE, size: 0, color: 'FFFFFF' };
const noBorders = { top: NONE, bottom: NONE, left: NONE, right: NONE };
const line = (size = 4) => ({ style: BorderStyle.SINGLE, size, color: BLACK });
const boxed = (size = 2) => ({ style: BorderStyle.SINGLE, size, color: BLACK, space: 8 });

/* ---------- helpers ---------- */
const run = (text, o = {}) => new TextRun({
  text,
  font: FONT,
  size: o.size || 20,
  bold: !!o.bold,
  italics: !!o.italics,
  color: BLACK,
  allCaps: !!o.caps,
  characterSpacing: o.track || 0,
});

/* a field for the school to replace — bold, so it is easy to find */
const fill = (label, o = {}) => new TextRun({
  text: '[ ' + label + ' ]',
  font: FONT,
  size: o.size || 20,
  bold: true,
  color: BLACK,
});

const p = (children, o = {}) => new Paragraph({
  children: Array.isArray(children) ? children : [children],
  alignment: o.align,
  spacing: { before: o.before || 0, after: o.after === undefined ? 160 : o.after, line: o.line || 280 },
  indent: o.indent,
  border: o.border,
  keepNext: o.keepNext,
});

const text = (s, o = {}) => p(run(s, o), o);
const gap = (h) => new Paragraph({ children: [], spacing: { before: 0, after: h, line: 240 } });

const cell = (children, o = {}) => new TableCell({
  children: Array.isArray(children) ? children : [children],
  width: { size: o.width, type: WidthType.DXA },
  columnSpan: o.span,
  borders: o.borders || noBorders,
  margins: {
    top: o.mt === undefined ? 40 : o.mt,
    bottom: o.mb === undefined ? 40 : o.mb,
    left: o.ml === undefined ? 0 : o.ml,
    right: o.mr === undefined ? 0 : o.mr,
  },
  verticalAlign: o.valign,
});

const table = (rows, widths) => new Table({
  rows, columnWidths: widths,
  width: { size: widths.reduce((a, b) => a + b, 0), type: WidthType.DXA },
  borders: noBorders,
});

/* numbered clause with a bold lead-in */
const clause = (n, lead, children) => new Paragraph({
  spacing: { after: 200, line: 290 },
  children: [run(n + '.  ' + lead + ' ', { bold: true }), ...children],
});

/* small heading */
const heading = (s, o = {}) => p(run(s, { bold: true, caps: true, track: 14, size: 18 }),
  { after: o.after === undefined ? 120 : o.after, keepNext: true, line: 260 });

/* ---------- letterhead ---------- */
const headCols = [1800, 7946];
const letterhead = table([
  new TableRow({
    children: [
      cell([
        p(run('SCHOOL', { size: 15, track: 10 }), { align: AlignmentType.CENTER, after: 0, line: 200 }),
        p(run('CREST', { size: 15, track: 10 }), { align: AlignmentType.CENTER, after: 0, line: 200 }),
        p(fill('logo', { size: 15 }), { align: AlignmentType.CENTER, after: 0, line: 200 }),
      ], { width: headCols[0], valign: VerticalAlign.CENTER, mr: 220 }),
      cell([
        p(run('Maputo International College', { size: 27, bold: true, track: 4 }), { after: 60, line: 260 }),
        p([fill('street address', { size: 17 }), run('  ·  ', { size: 17 }), fill('city, postal code', { size: 17 }), run('  ·  ', { size: 17 }), fill('country', { size: 17 })], { after: 40, line: 240 }),
        p([
          run('Tel ', { size: 17 }), fill('telephone', { size: 17 }),
          run('   Email ', { size: 17 }), fill('email', { size: 17 }),
          run('   Cambridge centre ', { size: 17 }), run('MZ040', { size: 17, bold: true }),
        ], { after: 0, line: 240 }),
      ], { width: headCols[1], valign: VerticalAlign.CENTER }),
    ],
  }),
], headCols);

/* ---------- signature block ---------- */
const sigCols = [5600, 4146];
const signature = table([
  new TableRow({
    children: [
      cell([
        gap(580),
        new Paragraph({ children: [], spacing: { after: 70 }, border: { bottom: line(6) } }),
        p(fill("Principal's full name"), { after: 50, line: 250 }),
        p(run('Principal', { bold: true }), { after: 30, line: 250 }),
        p(run('Maputo International College'), { after: 0, line: 250 }),
      ], { width: sigCols[0], mr: 320 }),
      cell([
        gap(140),
        p(run('OFFICIAL SCHOOL STAMP', { size: 15, track: 10 }), { align: AlignmentType.CENTER, after: 50, line: 220 }),
        p(fill('affix here', { size: 15 }), { align: AlignmentType.CENTER, after: 0, line: 220 }),
        gap(340),
      ], {
        width: sigCols[1], valign: VerticalAlign.CENTER,
        borders: { top: line(4), bottom: line(4), left: line(4), right: line(4) },
        ml: 130, mr: 130, mt: 130, mb: 130,
      }),
    ],
  }),
], sigCols);

/* ---------- guidance table (page 2) ---------- */
const noteRow = (field, guidance) => new TableRow({
  children: [
    cell(p(run(field, { size: 18, bold: true }), { after: 0, line: 250 }),
      { width: 3100, borders: { top: NONE, bottom: line(2), left: NONE, right: NONE }, mt: 80, mb: 80, mr: 220 }),
    cell(p(run(guidance, { size: 18 }), { after: 0, line: 250 }),
      { width: 6646, borders: { top: NONE, bottom: line(2), left: NONE, right: NONE }, mt: 80, mb: 80 }),
  ],
});

const pageFooter = new Footer({
  children: [
    new Paragraph({
      spacing: { before: 0, after: 0, line: 240 },
      border: { top: line(2) },
      alignment: AlignmentType.CENTER,
      children: [run('Draft — for completion by the issuing school. Not valid until completed, signed, dated and stamped.', { size: 15 })],
    }),
  ],
});

/* ---------- document ---------- */
const doc = new Document({
  creator: 'Office of the Principal',
  title: 'Confirmation of Academic Standing',
  description: 'Draft principal’s letter for completion by the issuing school',
  styles: { default: { document: { run: { font: FONT, size: 20, color: BLACK } } } },
  sections: [{
    properties: {
      page: {
        size: { width: PAGE_W, height: PAGE_H },
        margin: { top: 1000, right: MARGIN, bottom: 900, left: MARGIN, footer: 400 },
      },
    },
    footers: { default: pageFooter },
    children: [
      /* ===== PAGE 1 — THE LETTER ===== */
      letterhead,
      new Paragraph({ children: [], spacing: { before: 80, after: 260 }, border: { bottom: line(6) } }),

      table([
        new TableRow({
          children: [
            cell(p([run('Ref. '), fill('reference number')], { after: 0, line: 260 }), { width: 4873 }),
            cell(p([run('Date  '), fill('date of issue')], { align: AlignmentType.RIGHT, after: 0, line: 260 }), { width: 4873 }),
          ],
        }),
      ], [4873, 4873]),
      gap(300),

      p(run('To whom it may concern', { bold: true, caps: true, track: 14 }), { after: 260 }),

      p(run('Confirmation of Academic Standing', { size: 24, bold: true }), { after: 70, line: 260 }),
      new Paragraph({
        spacing: { after: 280, line: 260 },
        border: { bottom: line(2) },
        children: [run('Xavier Keylane Arianne Vasconcelos  ·  Date of birth 02 February 2007', { size: 18 })],
      }),

      p(run('I write in my capacity as Principal of Maputo International College to confirm the academic standing of the above-named student.'), { after: 240 }),

      clause(1, 'Candidature.', [
        run('Xavier Keylane Arianne Vasconcelos was a registered candidate of this College under Cambridge International centre number MZ040. The student sat Cambridge IGCSE examinations in the June 2024 and November 2024 series, and Cambridge International AS Level examinations in the November 2025 series, under candidate numbers MZ040/0036, MZ040/0016 and MZ040/0062 respectively.'),
      ]),

      clause(2, 'Examination record.', [
        run('Eleven syllabuses were reported across the three series: seven at Cambridge IGCSE and four at Cambridge International AS Level. The full record appears on the student’s Cambridge certificates and statements of results, and is summarised in the College’s Certificate of Cumulative Grade Point Average issued separately.'),
      ]),

      clause(3, 'Class standing.', [
        run('On the College’s internal ranking of the '), fill('cohort, e.g. November 2025 AS Level'),
        run(' cohort, comprising '), fill('number'),
        run(' students, this student was placed '), fill('position'),
        run(', which falls within the top '), fill('percentage'),
        run('% of the cohort. The ranking is calculated on '), fill('basis of ranking'),
        run(' and covers the period '), fill('period covered'), run('.'),
      ]),

      clause(4, 'Remarks.', [
        fill('optional: conduct, contribution to College life, particular strengths, or delete this clause'),
      ]),

      p([
        run('This confirmation is issued at the request of the student for the purpose of '),
        fill('purpose, e.g. university application'),
        run('. It is issued on the authority of this College and should be read together with the student’s Cambridge award documents. Any queries regarding its content may be addressed to the Office of the Principal at '),
        fill('contact details'), run('.'),
      ], { after: 300 }),

      p(run('Yours faithfully,'), { after: 0 }),

      signature,

      /* ===== PAGE 2 — NOTES FOR THE ISSUING OFFICE ===== */
      new Paragraph({ children: [new PageBreak()], spacing: { after: 0 } }),

      p(run('Notes for the issuing office', { size: 24, bold: true }), { after: 70, line: 260 }),
      new Paragraph({
        spacing: { after: 260, line: 260 },
        border: { bottom: line(2) },
        children: [run('This page is guidance only. Delete it before the letter is issued.', { size: 18 })],
      }),

      /* a one-cell table, not a four-sided paragraph border: docx-js emits
         pBdr children as top/bottom/left/right, which the schema rejects */
      table([
        new TableRow({
          children: [
            cell(p([
              run('The class-standing figures in clause 3 have deliberately been left blank. ', { bold: true }),
              run('They must be taken from the College’s own ranking records and verified before the letter is signed. Nothing in the student’s Cambridge results establishes a cohort position on its own, and a percentile stated without a ranking behind it is not a claim the College can stand behind.'),
            ], { after: 0, line: 280 }), {
              width: W,
              borders: { top: line(2), bottom: line(2), left: line(2), right: line(2) },
              mt: 150, mb: 150, ml: 170, mr: 170,
            }),
          ],
        }),
      ], [W]),
      gap(280),

      heading('Fields to complete'),
      table([
        noteRow('Letterhead', 'Insert the College crest, street address, telephone and email. The Cambridge centre number MZ040 is already filled in.'),
        noteRow('Ref. number', 'The College’s own correspondence reference, so the letter can be traced if a receiving institution writes back.'),
        noteRow('Date of issue', 'The date of signature. Receiving institutions commonly disregard confirmations older than twelve months.'),
        noteRow('Cohort', 'Name the group ranked against — for example the November 2025 AS Level cohort. A percentile means nothing without the group it refers to.'),
        noteRow('Number of students', 'The size of that cohort. A top-20% claim in a cohort of ten students is a different statement from one in a cohort of two hundred.'),
        noteRow('Position', 'Either a rank (“12th”) or a band (“in the top quintile”), whichever the College’s records actually support.'),
        noteRow('Percentage', 'Must follow arithmetically from the position and the cohort size.'),
        noteRow('Basis of ranking', 'State what the ranking is computed from — for example aggregate Cambridge grade points, internal assessment averages, or a weighted combination. Receiving institutions ask.'),
        noteRow('Period covered', 'The academic years or examination series the ranking spans.'),
        noteRow('Remarks', 'Optional. Delete clause 4 entirely if the College does not wish to add remarks.'),
        noteRow('Purpose', 'What the student requested the letter for.'),
        noteRow('Contact details', 'A verification route — normally the Registrar’s or Principal’s office email and telephone.'),
        noteRow('Signature block', 'The Principal’s name, wet signature and the College stamp. Unsigned and unstamped, the letter carries no weight.'),
      ], [3100, 6646]),
      gap(280),

      heading('Before signing'),
      p(run('Every field for completion is shown in bold inside square brackets. Search the document for an opening square bracket to step through them, and check that none remain once the letter is filled in. The draft notice at the foot of each page must also be deleted — it sits in the page footer, so double-click near the foot of the page to edit it.'), { after: 200 }),
      p(run('A letter confirming class standing is a statement of fact by the College about its own records. It should be issued only where those records exist and support what clause 3 says.', { italics: true }), { after: 0 }),
    ],
  }],
});

Packer.toBuffer(doc).then((buf) => {
  fs.writeFileSync(process.argv[2] || 'letter.docx', buf);
  console.log('written', process.argv[2], buf.length, 'bytes');
});
