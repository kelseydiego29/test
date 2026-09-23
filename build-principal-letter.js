const fs = require('fs');
const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell, WidthType,
  AlignmentType, BorderStyle, ShadingType, PageBreak, Footer, PositionalTab,
  PositionalTabAlignment, PositionalTabRelativeTo, PositionalTabLeader, VerticalAlign,
} = require('docx');

/* ---------- tokens ---------- */
const ACCENT = '0F3D5C';
const ACCENT_SOFT = 'EEF3F7';
const GOLD = '9A7B3F';
const INK = '1C2230';
const INK_SOFT = '3A4354';
const INK_FAINT = '7B8496';
const FILL_BLUE = '1F5F8B';      // colour of the fields to be completed
const RULE = 'C9D0DC';
const RULE_SOFT = 'E4E8EF';

const SERIF = 'Georgia';
const SANS = 'Calibri';

/* A4, letter margins */
const PAGE_W = 11906, PAGE_H = 16838, MARGIN = 1080;
const W = PAGE_W - MARGIN * 2; // 9746

const NONE = { style: BorderStyle.NONE, size: 0, color: 'FFFFFF' };
const noBorders = { top: NONE, bottom: NONE, left: NONE, right: NONE };
const line = (color, size = 4) => ({ style: BorderStyle.SINGLE, size, color });

/* ---------- helpers ---------- */
const run = (text, o = {}) => new TextRun({
  text,
  font: o.font || SANS,
  size: o.size || 21,
  bold: !!o.bold,
  italics: !!o.italics,
  color: o.color || INK,
  allCaps: !!o.caps,
  characterSpacing: o.track || 0,
});

/* a field the school replaces */
const fill = (label) => new TextRun({
  text: '[ ' + label + ' ]',
  font: SANS,
  size: 21,
  bold: true,
  color: FILL_BLUE,
});

const p = (children, o = {}) => new Paragraph({
  children: Array.isArray(children) ? children : [children],
  alignment: o.align,
  spacing: { before: o.before || 0, after: o.after === undefined ? 160 : o.after, line: o.line || 276 },
  indent: o.indent,
  border: o.border,
  shading: o.fill ? { type: ShadingType.CLEAR, fill: o.fill, color: 'auto' } : undefined,
  keepNext: o.keepNext,
});

const text = (s, o = {}) => p(run(s, o), o);
const gap = (h) => new Paragraph({ children: [], spacing: { before: 0, after: h, line: 240 } });

const cell = (children, o = {}) => new TableCell({
  children: Array.isArray(children) ? children : [children],
  width: { size: o.width, type: WidthType.DXA },
  columnSpan: o.span,
  shading: o.fill ? { type: ShadingType.CLEAR, fill: o.fill, color: 'auto' } : undefined,
  borders: o.borders || noBorders,
  margins: { top: o.mt === undefined ? 40 : o.mt, bottom: o.mb === undefined ? 40 : o.mb, left: o.ml === undefined ? 0 : o.ml, right: o.mr === undefined ? 0 : o.mr },
  verticalAlign: o.valign,
});

const table = (rows, widths) => new Table({
  rows, columnWidths: widths,
  width: { size: widths.reduce((a, b) => a + b, 0), type: WidthType.DXA },
  borders: noBorders,
});

/* numbered clause: bold lead-in, then body runs */
const clause = (n, lead, children) => new Paragraph({
  spacing: { after: 200, line: 288 },
  children: [
    run(n + '.  ', { bold: true, color: ACCENT }),
    run(lead + ' ', { bold: true }),
    ...children,
  ],
});

/* ---------- letterhead ---------- */
const headCols = [1900, 7846];
const letterhead = table([
  new TableRow({
    children: [
      cell([
        p(run('SCHOOL', { size: 15, color: INK_FAINT, track: 10 }), { align: AlignmentType.CENTER, after: 0, line: 200 }),
        p(run('CREST', { size: 15, color: INK_FAINT, track: 10 }), { align: AlignmentType.CENTER, after: 0, line: 200 }),
        p(run('[ insert logo ]', { size: 14, color: FILL_BLUE }), { align: AlignmentType.CENTER, after: 0, line: 200 }),
      ], { width: headCols[0], valign: VerticalAlign.CENTER, mr: 200 }),
      cell([
        p(run('Maputo International College', { font: SERIF, size: 30, bold: true, color: ACCENT, track: 6 }), { after: 40, line: 240 }),
        p([fill('street address'), run('  ·  ', { size: 19, color: INK_FAINT }), fill('city, postal code'), run('  ·  ', { size: 19, color: INK_FAINT }), fill('country')], { after: 30, line: 240 }),
        p([run('Tel ', { size: 19, color: INK_FAINT }), fill('telephone'), run('   Email ', { size: 19, color: INK_FAINT }), fill('email'), run('   Cambridge centre ', { size: 19, color: INK_FAINT }), run('MZ040', { size: 19, bold: true })], { after: 0, line: 240 }),
      ], { width: headCols[1], valign: VerticalAlign.CENTER }),
    ],
  }),
], headCols);

/* ---------- signature block ---------- */
const sigCols = [5400, 4346];
const signature = table([
  new TableRow({
    children: [
      cell([
        gap(560),
        new Paragraph({ children: [], spacing: { after: 60 }, border: { bottom: line(INK_SOFT, 6) } }),
        p([fill("Principal's full name")], { after: 40, line: 240 }),
        p(run('Principal', { bold: true, size: 20 }), { after: 20, line: 240 }),
        p(run('Maputo International College', { size: 20, color: INK_SOFT }), { after: 0, line: 240 }),
      ], { width: sigCols[0], mr: 300 }),
      cell([
        gap(120),
        p(run('OFFICIAL SCHOOL STAMP', { size: 15, color: INK_FAINT, track: 10 }), { align: AlignmentType.CENTER, after: 40, line: 220 }),
        p(run('[ affix here ]', { size: 15, color: FILL_BLUE }), { align: AlignmentType.CENTER, after: 0, line: 220 }),
        gap(320),
      ], {
        width: sigCols[1], valign: VerticalAlign.CENTER,
        borders: { top: line(RULE, 4), bottom: line(RULE, 4), left: line(RULE, 4), right: line(RULE, 4) },
        ml: 120, mr: 120, mt: 120, mb: 120,
      }),
    ],
  }),
], sigCols);

/* ---------- completion notes (page 2) ---------- */
const noteRow = (field, guidance) => new TableRow({
  children: [
    cell(p(run(field, { size: 19, bold: true }), { after: 0, line: 240 }),
      { width: 3200, borders: { top: NONE, bottom: line(RULE_SOFT, 4), left: NONE, right: NONE }, mt: 70, mb: 70, mr: 200 }),
    cell(p(run(guidance, { size: 19, color: INK_SOFT }), { after: 0, line: 240 }),
      { width: 6546, borders: { top: NONE, bottom: line(RULE_SOFT, 4), left: NONE, right: NONE }, mt: 70, mb: 70 }),
  ],
});

const pageFooter = new Footer({
  children: [
    new Paragraph({
      spacing: { before: 0, after: 0, line: 240 },
      border: { top: line(RULE_SOFT, 4) },
      children: [
        run('DRAFT — FOR COMPLETION BY THE ISSUING SCHOOL', { size: 14, bold: true, color: GOLD, track: 14 }),
        new TextRun({
          children: [new PositionalTab({
            alignment: PositionalTabAlignment.RIGHT,
            relativeTo: PositionalTabRelativeTo.MARGIN,
            leader: PositionalTabLeader.NONE,
          })],
        }),
        run('Not valid until completed, signed, dated and stamped', { size: 14, color: INK_FAINT }),
      ],
    }),
  ],
});

/* ---------- document ---------- */
const doc = new Document({
  creator: 'Office of the Principal',
  title: 'Confirmation of Academic Standing',
  description: 'Draft principal’s letter for completion by the issuing school',
  styles: { default: { document: { run: { font: SANS, size: 21, color: INK } } } },
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
      new Paragraph({ children: [], spacing: { before: 60, after: 240 }, border: { bottom: line(ACCENT, 12) } }),

      table([
        new TableRow({
          children: [
            cell(p([run('Ref. ', { size: 20, color: INK_FAINT }), fill('reference number')], { after: 0, line: 260 }),
              { width: 4873 }),
            cell(p([run('Date  ', { size: 20, color: INK_FAINT }), fill('date of issue')], { align: AlignmentType.RIGHT, after: 0, line: 260 }),
              { width: 4873 }),
          ],
        }),
      ], [4873, 4873]),
      gap(300),

      text('To whom it may concern', { bold: true, caps: true, track: 16, size: 21, after: 260 }),

      new Paragraph({
        spacing: { after: 60, line: 260 },
        children: [run('Confirmation of Academic Standing', { font: SERIF, size: 26, bold: true, color: ACCENT })],
      }),
      new Paragraph({
        spacing: { after: 280, line: 260 },
        border: { bottom: line(RULE, 4) },
        children: [run('Xavier Keylane Arianne Vasconcelos  ·  Date of birth 02 February 2007', { size: 20, color: INK_SOFT })],
      }),

      p([run('I write in my capacity as Principal of Maputo International College to confirm the academic standing of the above-named student.', {})], { after: 240 }),

      clause(1, 'Candidature.', [
        run('Xavier Keylane Arianne Vasconcelos was a registered candidate of this College under Cambridge International centre number MZ040. The student sat Cambridge IGCSE examinations in the June 2024 and November 2024 series, and Cambridge International AS Level examinations in the November 2025 series, under candidate numbers MZ040/0036, MZ040/0016 and MZ040/0062 respectively.', { color: INK_SOFT }),
      ]),

      clause(2, 'Examination record.', [
        run('Eleven syllabuses were reported across the three series: seven at Cambridge IGCSE and four at Cambridge International AS Level. The full record appears on the student’s Cambridge certificates and statements of results, and is summarised in the College’s Certificate of Cumulative Grade Point Average issued separately.', { color: INK_SOFT }),
      ]),

      clause(3, 'Class standing.', [
        run('On the College’s internal ranking of the ', { color: INK_SOFT }),
        fill('cohort, e.g. November 2025 AS Level'),
        run(' cohort, comprising ', { color: INK_SOFT }),
        fill('number'),
        run(' students, this student was placed ', { color: INK_SOFT }),
        fill('position'),
        run(', which falls within the top ', { color: INK_SOFT }),
        fill('percentage'),
        run('% of the cohort. The ranking is calculated on ', { color: INK_SOFT }),
        fill('basis of ranking'),
        run(' and covers the period ', { color: INK_SOFT }),
        fill('period covered'),
        run('.', { color: INK_SOFT }),
      ]),

      clause(4, 'Remarks.', [
        fill('optional: conduct, contribution to College life, particular strengths, or delete this clause'),
      ]),

      p([
        run('This confirmation is issued at the request of the student for the purpose of ', { color: INK_SOFT }),
        fill('purpose, e.g. university application'),
        run('. It is issued on the authority of this College and should be read together with the student’s Cambridge award documents. Any queries regarding its content may be addressed to the Office of the Principal at ', { color: INK_SOFT }),
        fill('contact details'),
        run('.', { color: INK_SOFT }),
      ], { after: 300 }),

      text('Yours faithfully,', { after: 0 }),

      signature,

      /* ===== PAGE 2 — NOTES FOR THE ISSUING OFFICE ===== */
      new Paragraph({ children: [new PageBreak()], spacing: { after: 0 } }),

      new Paragraph({
        spacing: { after: 60, line: 260 },
        children: [run('Notes for the issuing office', { font: SERIF, size: 26, bold: true, color: ACCENT })],
      }),
      new Paragraph({
        spacing: { after: 240, line: 260 },
        border: { bottom: line(RULE, 4) },
        children: [run('This page is guidance only. Delete it before the letter is issued.', { size: 20, color: INK_SOFT })],
      }),

      new Paragraph({
        spacing: { before: 40, after: 260, line: 276 },
        shading: { type: ShadingType.CLEAR, fill: ACCENT_SOFT, color: 'auto' },
        border: { left: line(ACCENT, 18) },
        indent: { left: 180, right: 180 },
        children: [
          run('The class-standing figures in clause 3 have deliberately been left blank. ', { bold: true, color: ACCENT }),
          run('They must be taken from the College’s own ranking records and verified before the letter is signed. Nothing in the student’s Cambridge results establishes a cohort position on its own, and a percentile stated without a ranking behind it is not a claim the College can stand behind.', { color: INK_SOFT }),
        ],
      }),

      text('Fields to complete', { bold: true, caps: true, track: 14, size: 18, color: ACCENT, after: 100 }),
      table([
        noteRow('Letterhead', 'Insert the College crest, street address, telephone and email. The Cambridge centre number MZ040 is already filled in.'),
        noteRow('Ref. number', 'The College’s own correspondence reference, so the letter can be traced if a receiving institution writes back.'),
        noteRow('Date of issue', 'The date of signature. Receiving institutions commonly disregard confirmations older than twelve months.'),
        noteRow('Cohort', 'Name the group ranked against — for example the November 2025 AS Level cohort. A percentile means nothing without the group it refers to.'),
        noteRow('Number of students', 'The size of that cohort. A top-20% claim in a cohort of ten students is a different statement from one in a cohort of two hundred.'),
        noteRow('Position', 'Either a rank ("12th") or a band ("in the top quintile"), whichever the College’s records actually support.'),
        noteRow('Percentage', 'Must follow arithmetically from the position and the cohort size.'),
        noteRow('Basis of ranking', 'State what the ranking is computed from — for example aggregate Cambridge grade points, internal assessment averages, or a weighted combination. Receiving institutions ask.'),
        noteRow('Period covered', 'The academic years or examination series the ranking spans.'),
        noteRow('Remarks', 'Optional. Delete clause 4 entirely if the College does not wish to add remarks.'),
        noteRow('Purpose', 'What the student requested the letter for.'),
        noteRow('Contact details', 'A verification route — normally the Registrar’s or Principal’s office email and telephone.'),
        noteRow('Signature block', 'The Principal’s name, wet signature and the College stamp. Unsigned and unstamped, the letter carries no weight.'),
      ], [3200, 6546]),
      gap(260),

      text('Before signing', { bold: true, caps: true, track: 14, size: 18, color: ACCENT, after: 100 }),
      p([run('Check that every blue bracketed field has been replaced, that no bracket characters remain in the text, and that the footer marking this document a draft has been deleted. The footer is in the page footer area — double-click near the foot of the page to edit it.', { color: INK_SOFT })], { after: 200 }),
      p([run('A letter confirming class standing is a statement of fact by the College about its own records. It should be issued only where those records exist and support what clause 3 says.', { color: INK_SOFT, italics: true })], { after: 0 }),
    ],
  }],
});

Packer.toBuffer(doc).then((buf) => {
  fs.writeFileSync(process.argv[2] || 'letter.docx', buf);
  console.log('written', process.argv[2], buf.length, 'bytes');
});
