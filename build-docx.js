const fs = require('fs');
const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell, WidthType,
  AlignmentType, BorderStyle, ShadingType, PageBreak, Footer, PositionalTab,
  PositionalTabAlignment, PositionalTabRelativeTo, PositionalTabLeader,
  VerticalAlign, HeadingLevel,
} = require('docx');

/* ---------- tokens ---------- */
const ACCENT = '0F3D5C';
const ACCENT_SOFT = 'EEF3F7';
const GOLD = '9A7B3F';
const INK = '1C2230';
const INK_SOFT = '4A5468';
const INK_FAINT = '7B8496';
const RULE = 'C9D0DC';
const RULE_SOFT = 'E4E8EF';
const WHITE = 'FFFFFF';

const SERIF = 'Georgia';
const SANS = 'Calibri';

/* A4 with ~14mm side margins */
const PAGE_W = 11906, PAGE_H = 16838, MARGIN = 800;
const W = PAGE_W - MARGIN * 2; // 10306

const NONE = { style: BorderStyle.NONE, size: 0, color: 'FFFFFF' };
const noBorders = { top: NONE, bottom: NONE, left: NONE, right: NONE };
const line = (color, size = 4) => ({ style: BorderStyle.SINGLE, size, color });

/* ---------- small helpers ---------- */
const run = (text, o = {}) => new TextRun({
  text,
  font: o.font || SANS,
  size: o.size || 19,            // half-points
  bold: !!o.bold,
  color: o.color || INK,
  allCaps: !!o.caps,
  characterSpacing: o.track || 0,
  italics: !!o.italics,
});

const p = (children, o = {}) => new Paragraph({
  children: Array.isArray(children) ? children : [children],
  alignment: o.align,
  spacing: { before: o.before || 0, after: o.after === undefined ? 60 : o.after, line: o.line },
  border: o.border,
  shading: o.fill ? { type: ShadingType.CLEAR, fill: o.fill, color: 'auto' } : undefined,
  indent: o.indent,
  keepNext: o.keepNext,
});

const text = (s, o = {}) => p(run(s, o), o);

const cell = (children, o = {}) => new TableCell({
  children: Array.isArray(children) ? children : [children],
  width: { size: o.width, type: WidthType.DXA },
  columnSpan: o.span,
  shading: o.fill ? { type: ShadingType.CLEAR, fill: o.fill, color: 'auto' } : undefined,
  borders: o.borders || noBorders,
  margins: { top: o.mt === undefined ? 40 : o.mt, bottom: o.mb === undefined ? 40 : o.mb, left: o.ml === undefined ? 100 : o.ml, right: 100 },
  verticalAlign: o.valign,
});

const table = (rows, widths) => new Table({
  rows,
  columnWidths: widths,
  width: { size: widths.reduce((a, b) => a + b, 0), type: WidthType.DXA },
  borders: noBorders,
});

/* rule as a bottom-bordered empty paragraph */
const hr = (color = RULE, size = 6, after = 120) => new Paragraph({
  children: [],
  spacing: { before: 0, after },
  border: { bottom: line(color, size) },
});

const gap = (h) => new Paragraph({ children: [], spacing: { before: 0, after: h } });

/* ---------- grade tables ---------- */
const GRADE_COLS = [1150, 4150, 1300, 900, 1200, 1606]; // = 10306

function gradeHeader() {
  const th = (s, align) => cell(
    p(run(s, { size: 13, bold: true, color: INK_FAINT, caps: true, track: 12 }), { align, after: 0 }),
    { width: 0, borders: { top: line(RULE, 6), bottom: line(RULE, 6), left: NONE, right: NONE }, mt: 50, mb: 50 }
  );
  const labels = [
    ['Syllabus', AlignmentType.LEFT], ['Subject', AlignmentType.LEFT], ['Series', AlignmentType.LEFT],
    ['Grade', AlignmentType.CENTER], ['PUM', AlignmentType.RIGHT], ['Grade pts', AlignmentType.RIGHT],
  ];
  return new TableRow({
    children: labels.map(([s, a], i) => {
      const c = th(s, a);
      c.options.width = { size: GRADE_COLS[i], type: WidthType.DXA };
      return c;
    }),
    tableHeader: true,
  });
}

function gradeRow(r) {
  const b = { top: NONE, bottom: line(RULE_SOFT, 4), left: NONE, right: NONE };
  return new TableRow({
    children: [
      cell(p(run(r.code, { size: 18, color: INK_FAINT }), { after: 0 }), { width: GRADE_COLS[0], borders: b }),
      cell(p(run(r.subject, { size: 19 }), { after: 0 }), { width: GRADE_COLS[1], borders: b }),
      cell(p(run(r.series, { size: 19 }), { after: 0 }), { width: GRADE_COLS[2], borders: b }),
      cell(p(run(r.grade, { size: 19, bold: true }), { align: AlignmentType.CENTER, after: 0 }), { width: GRADE_COLS[3], borders: b }),
      cell(p(run(r.pum, { size: 19 }), { align: AlignmentType.RIGHT, after: 0 }), { width: GRADE_COLS[4], borders: b }),
      cell(p(run(r.pts, { size: 19 }), { align: AlignmentType.RIGHT, after: 0 }), { width: GRADE_COLS[5], borders: b }),
    ],
  });
}

function gradeFoot(label, pum, pts) {
  const b = { top: line(RULE, 6), bottom: NONE, left: NONE, right: NONE };
  const f = (s, align, span) => cell(
    p(run(s, { size: 18, bold: true, color: ACCENT }), { align, after: 0 }),
    { width: 0, span, fill: ACCENT_SOFT, borders: b }
  );
  const c0 = f(label, AlignmentType.LEFT, 3);
  c0.options.width = { size: GRADE_COLS[0] + GRADE_COLS[1] + GRADE_COLS[2], type: WidthType.DXA };
  const c1 = f('—', AlignmentType.CENTER); c1.options.width = { size: GRADE_COLS[3], type: WidthType.DXA };
  const c2 = f(pum, AlignmentType.RIGHT); c2.options.width = { size: GRADE_COLS[4], type: WidthType.DXA };
  const c3 = f(pts, AlignmentType.RIGHT); c3.options.width = { size: GRADE_COLS[5], type: WidthType.DXA };
  return new TableRow({ children: [c0, c1, c2, c3] });
}

/* section heading with right-aligned note on the same line */
const sectionHead = (title, note) => new Paragraph({
  spacing: { before: 120, after: 60 },
  keepNext: true,
  children: [
    run(title, { size: 18, bold: true, color: ACCENT, caps: true, track: 14 }),
    ...(note ? [new TextRun({
      children: [new PositionalTab({
        alignment: PositionalTabAlignment.RIGHT,
        relativeTo: PositionalTabRelativeTo.MARGIN,
        leader: PositionalTabLeader.NONE,
      })],
      font: SANS, size: 17, color: INK_FAINT,
    }), run(note, { size: 17, color: INK_FAINT })] : []),
  ],
});

/* ---------- data ---------- */
const IGCSE = [
  { code: '0450', subject: 'Business Studies', series: 'Jun 2024', grade: 'A', pum: '80%', pts: '4.00' },
  { code: '0460', subject: 'Geography', series: 'Jun 2024', grade: 'B', pum: '75%', pts: '3.00' },
  { code: '0500', subject: 'First Language English', series: 'Jun 2024', grade: 'B', pum: '70%', pts: '3.00' },
  { code: '0580', subject: 'Mathematics', series: 'Nov 2024', grade: 'A', pum: '80%', pts: '4.00' },
  { code: '0452', subject: 'Accounting', series: 'Nov 2024', grade: 'B', pum: '77%', pts: '3.00' },
  { code: '0455', subject: 'Economics', series: 'Nov 2024', grade: 'C', pum: '68%', pts: '2.00' },
  { code: '0417', subject: 'Information and Communication Technology', series: 'Nov 2024', grade: 'C', pum: '64%', pts: '2.00' },
];
const ASL = [
  { code: '9706', subject: 'Accounting', series: 'Nov 2025', grade: 'c', pum: '67%', pts: '2.00' },
  { code: '9609', subject: 'Business', series: 'Nov 2025', grade: 'c', pum: '62%', pts: '2.00' },
  { code: '9708', subject: 'Economics', series: 'Nov 2025', grade: 'c', pum: '60%', pts: '2.00' },
  { code: '9709', subject: 'Mathematics', series: 'Nov 2025', grade: 'c', pum: '62%', pts: '2.00' },
];

/* ---------- page 1 blocks ---------- */
const issuerCols = [1500, 5400, 3406];
const issuerRow = new TableRow({
  children: [
    cell([
      p(run('SCHOOL CREST', { size: 13, color: INK_FAINT, track: 8 }), { after: 0 }),
      p(run('[ logo here ]', { size: 13, color: INK_FAINT }), { after: 0 }),
    ], { width: issuerCols[0], valign: VerticalAlign.CENTER }),
    cell([
      new Paragraph({
        spacing: { after: 20 },
        border: { bottom: line(RULE_SOFT, 4) },
        children: [run('[ Issuing institution name ]', { size: 20, color: INK_FAINT })],
      }),
      p(run('Office of the Registrar', { size: 16, color: INK_FAINT, caps: true, track: 12 }), { after: 0 }),
    ], { width: issuerCols[1], valign: VerticalAlign.CENTER }),
    cell([
      p([run('Ref. No. ', { size: 16, bold: true, color: INK_SOFT }), run('______________', { size: 16, color: INK_FAINT })], { align: AlignmentType.RIGHT, after: 20 }),
      p([run('Date of issue ', { size: 16, bold: true, color: INK_SOFT }), run('______________', { size: 16, color: INK_FAINT })], { align: AlignmentType.RIGHT, after: 0 }),
    ], { width: issuerCols[2], valign: VerticalAlign.CENTER }),
  ],
});

const PARTICULARS = [
  ['Date of birth', '02 February 2007'],
  ['Centre', 'MZ040 — Maputo Int. College'],
  ['Examination series', 'Jun 2024 – Nov 2025'],
  ['Syllabuses reported', '11'],
];
const partCols = [2200, 3100, 2600, 2406];
const particularsTable = table([
  new TableRow({
    children: PARTICULARS.map(([k, v], i) => cell([
      p(run(k, { size: 13, color: INK_FAINT, caps: true, track: 12 }), { after: 10 }),
      p(run(v, { size: 18, bold: true }), { after: 0 }),
    ], {
      width: partCols[i],
      fill: 'FBFCFD',
      borders: {
        top: line(RULE_SOFT, 4), bottom: line(RULE_SOFT, 4),
        left: i === 0 ? line(RULE_SOFT, 4) : NONE,
        right: line(RULE_SOFT, 4),
      },
      mt: 60, mb: 60,
    })),
  }),
], partCols);

const resCols = [2900, 2900, 4506];
const resultPanel = table([
  new TableRow({
    children: [
      cell([
        p(run('IGCSE stage GPA', { size: 13, color: INK_FAINT, caps: true, track: 12 }), { after: 30 }),
        p(run('3.00', { font: SERIF, size: 40, bold: true, color: INK }), { after: 30 }),
        p(run('21.00 points / 7 subjects', { size: 15, color: INK_FAINT }), { after: 0 }),
      ], {
        width: resCols[0], mt: 90, mb: 90,
        borders: { top: line(ACCENT, 12), bottom: line(ACCENT, 12), left: line(ACCENT, 12), right: line(RULE, 4) },
      }),
      cell([
        p(run('AS Level stage GPA', { size: 13, color: INK_FAINT, caps: true, track: 12 }), { after: 30 }),
        p(run('2.00', { font: SERIF, size: 40, bold: true, color: INK }), { after: 30 }),
        p(run('8.00 points / 4 subjects', { size: 15, color: INK_FAINT }), { after: 0 }),
      ], {
        width: resCols[1], mt: 90, mb: 90,
        borders: { top: line(ACCENT, 12), bottom: line(ACCENT, 12), left: NONE, right: line(RULE, 4) },
      }),
      cell([
        p(run('Cumulative GPA', { size: 13, color: 'C3D2DE', caps: true, track: 12 }), { after: 30 }),
        p(run('2.64', { font: SERIF, size: 56, bold: true, color: WHITE }), { after: 30 }),
        p(run('29.00 points / 11 subjects  ·  on a 4.00 scale', { size: 15, color: 'C3D2DE' }), { after: 0 }),
      ], {
        width: resCols[2], fill: ACCENT, mt: 90, mb: 90,
        borders: { top: line(ACCENT, 12), bottom: line(ACCENT, 12), left: NONE, right: line(ACCENT, 12) },
      }),
    ],
  }),
], resCols);

function signatureBlock(labels) {
  const n = labels.length;
  const w = Math.floor(W / n);
  const cols = labels.map((_, i) => (i === n - 1 ? W - w * (n - 1) : w));
  return table([
    new TableRow({
      children: cols.map((cw) => cell(gap(0), {
        width: cw, mt: 280, mb: 0,
        borders: { top: NONE, bottom: line(INK_SOFT, 6), left: NONE, right: NONE },
      })),
    }),
    new TableRow({
      children: cols.map((cw, i) => cell([
        p(run(labels[i][0], { size: 16, bold: true, color: INK_SOFT, caps: true, track: 10 }), { align: AlignmentType.CENTER, after: 10 }),
        p(run(labels[i][1], { size: 14, color: INK_FAINT }), { align: AlignmentType.CENTER, after: 0 }),
      ], { width: cw, mt: 60, mb: 0 })),
    }),
  ], cols);
}

/* ---------- page 2 tables ---------- */
const keyCols = [1700, 1500, 1200];
const KEY = [
  ['A* / a*', '4.00', '0'], ['A / a', '4.00', '2'], ['B / b', '3.00', '3'],
  ['C / c', '2.00', '6'], ['D / d', '1.00', '0'], ['E / e and below', '0.00', '0'],
];
const keyTable = table([
  new TableRow({
    tableHeader: true,
    children: ['Cambridge grade', 'Grade points', 'Count'].map((s, i) => cell(
      p(run(s, { size: 13, bold: true, color: INK_FAINT, caps: true, track: 12 }),
        { align: i === 0 ? AlignmentType.LEFT : AlignmentType.RIGHT, after: 0 }),
      { width: keyCols[i], borders: { top: line(RULE, 6), bottom: line(RULE, 6), left: NONE, right: NONE }, mt: 50, mb: 50 }
    )),
  }),
  ...KEY.map(([g, pts, ct]) => new TableRow({
    children: [g, pts, ct].map((s, i) => cell(
      p(run(s, { size: 19 }), { align: i === 0 ? AlignmentType.LEFT : AlignmentType.RIGHT, after: 0 }),
      { width: keyCols[i], borders: { top: NONE, bottom: line(RULE_SOFT, 4), left: NONE, right: NONE } }
    )),
  })),
  new TableRow({
    children: [
      cell(p(run('Total syllabuses converted', { size: 18, bold: true, color: ACCENT }), { after: 0 }),
        { width: keyCols[0] + keyCols[1], span: 2, fill: ACCENT_SOFT, borders: { top: line(RULE, 6), bottom: NONE, left: NONE, right: NONE } }),
      cell(p(run('11', { size: 18, bold: true, color: ACCENT }), { align: AlignmentType.RIGHT, after: 0 }),
        { width: keyCols[2], fill: ACCENT_SOFT, borders: { top: line(RULE, 6), bottom: NONE, left: NONE, right: NONE } }),
    ],
  }),
], keyCols);

const calcCols = [2700, 1300, 1600];
const CALC = [
  ['IGCSE grade points', '21.00', '÷ 7 = 3.00', false],
  ['AS Level grade points', '8.00', '÷ 4 = 2.00', false],
  ['Cumulative (unweighted)', '29.00', '÷ 11 = 2.64', true],
  ['Cumulative (credit-weighted)', '37.00', '÷ 15 = 2.47', false],
];
const calcTable = table(CALC.map(([a, b, c, bold]) => new TableRow({
  children: [a, b, c].map((s, i) => cell(
    p(run(s, { size: 18, bold, color: bold ? INK : (i === 0 ? INK_FAINT : INK_SOFT) }),
      { align: i === 0 ? AlignmentType.LEFT : AlignmentType.RIGHT, after: 0 }),
    { width: calcCols[i], borders: { top: NONE, bottom: line(RULE_SOFT, 4), left: NONE, right: NONE } }
  )),
})), calcCols);

const sideCols = [4500, 5806];
const keyCalcBlock = table([
  new TableRow({
    children: [
      cell([sectionHead('Grade conversion key'), keyTable], { width: sideCols[0], mt: 0, mb: 0, ml: 0 }),
      cell([sectionHead('Worked calculation'), calcTable], { width: sideCols[1], mt: 0, mb: 0 }),
    ],
  }),
], sideCols);

const srcCols = [3800, 1700, 1600, 1700, 1506];
const SRC = [
  ['International General Certificate of Secondary Education', 'June 2024', 'MZ040/0036', '0072782654', '3'],
  ['International General Certificate of Secondary Education', 'November 2024', 'MZ040/0016', '0073531190', '4'],
  ['General Certificate of Education (Advanced Subsidiary)', 'November 2025', 'MZ040/0062', '0075373702', '4'],
];
const srcTable = table([
  new TableRow({
    tableHeader: true,
    children: ['Qualification', 'Series', 'Candidate No.', 'Certificate No.', 'Syllabuses'].map((s, i) => cell(
      p(run(s, { size: 13, bold: true, color: INK_FAINT, caps: true, track: 12 }),
        { align: i === 4 ? AlignmentType.RIGHT : AlignmentType.LEFT, after: 0 }),
      { width: srcCols[i], borders: { top: line(RULE, 6), bottom: line(RULE, 6), left: NONE, right: NONE }, mt: 50, mb: 50 }
    )),
  }),
  ...SRC.map((r) => new TableRow({
    children: r.map((s, i) => cell(
      p(run(s, { size: 18, color: (i === 2 || i === 3) ? INK_FAINT : INK }),
        { align: i === 4 ? AlignmentType.RIGHT : AlignmentType.LEFT, after: 0 }),
      { width: srcCols[i], borders: { top: NONE, bottom: line(RULE_SOFT, 4), left: NONE, right: NONE } }
    )),
  })),
  new TableRow({
    children: [
      cell(p(run('Three award documents, with the corresponding Statements of Results', { size: 18, bold: true, color: ACCENT }), { after: 0 }),
        { width: srcCols[0] + srcCols[1] + srcCols[2] + srcCols[3], span: 4, fill: ACCENT_SOFT, borders: { top: line(RULE, 6), bottom: NONE, left: NONE, right: NONE } }),
      cell(p(run('11', { size: 18, bold: true, color: ACCENT }), { align: AlignmentType.RIGHT, after: 0 }),
        { width: srcCols[4], fill: ACCENT_SOFT, borders: { top: line(RULE, 6), bottom: NONE, left: NONE, right: NONE } }),
    ],
  }),
], srcCols);

const methodPara = (lead, body) => new Paragraph({
  spacing: { after: 90, line: 225 },
  children: [run(lead, { size: 17, bold: true }), run(' ' + body, { size: 17, color: INK_SOFT })],
});

/* ---------- footer ---------- */
const pageFooter = new Footer({
  children: [
    new Paragraph({
      spacing: { before: 0, after: 0 },
      border: { top: line(RULE_SOFT, 4) },
      children: [
        run('SPECIMEN — SAMPLE FORMAT', { size: 13, bold: true, color: GOLD, track: 16 }),
        new TextRun({
          children: [new PositionalTab({
            alignment: PositionalTabAlignment.RIGHT,
            relativeTo: PositionalTabRelativeTo.MARGIN,
            leader: PositionalTabLeader.NONE,
          })],
        }),
        run('Grade-conversion statement · not a Cambridge International award document', { size: 13, color: INK_FAINT }),
      ],
    }),
  ],
});

/* ---------- document ---------- */
const doc = new Document({
  creator: 'Office of the Registrar',
  title: 'Certificate of Cumulative Grade Point Average',
  description: 'Specimen CGPA conversion certificate',
  styles: {
    default: {
      document: { run: { font: SANS, size: 19, color: INK } },
    },
  },
  sections: [{
    properties: {
      page: {
        size: { width: PAGE_W, height: PAGE_H },
        margin: { top: 800, right: MARGIN, bottom: 700, left: MARGIN, footer: 360 },
      },
    },
    footers: { default: pageFooter },
    children: [
      /* ===== PAGE 1 ===== */
      table([issuerRow], issuerCols),
      hr(RULE, 6, 60),

      text('Certificate of', { align: AlignmentType.CENTER, font: SERIF, size: 38, bold: true, color: ACCENT, caps: true, track: 8, before: 60, after: 0 }),
      text('Cumulative Grade Point Average', { align: AlignmentType.CENTER, font: SERIF, size: 38, bold: true, color: ACCENT, caps: true, track: 8, after: 70 }),
      text('Conversion of Cambridge International Grades to a 4.00 Scale', { align: AlignmentType.CENTER, size: 15, color: INK_FAINT, caps: true, track: 22, after: 50 }),
      new Paragraph({ children: [], spacing: { after: 150 }, border: { bottom: line(GOLD, 4) }, indent: { left: 3400, right: 3400 } }),

      text('This is to certify that', { align: AlignmentType.CENTER, size: 19, color: INK_SOFT, after: 60 }),
      text('Xavier Keylane Arianne Vasconcelos', { align: AlignmentType.CENTER, font: SERIF, size: 28, bold: true, caps: true, track: 12, after: 70 }),
      new Paragraph({
        alignment: AlignmentType.CENTER,
        spacing: { after: 150, line: 250 },
        indent: { left: 700, right: 700 },
        children: [run('completed the Cambridge International examinations listed below at Maputo International College, and that the cumulative grade point average derived from those results is as stated in this certificate.', { size: 19, color: INK_SOFT })],
      }),

      particularsTable,
      gap(80),

      sectionHead('Part I — Cambridge IGCSE', 'Candidate Nos. MZ040/0036 & MZ040/0016'),
      table([gradeHeader(), ...IGCSE.map(gradeRow), gradeFoot('Seven syllabuses — IGCSE stage average', '73.4%', '3.00')], GRADE_COLS),
      gap(160),

      sectionHead('Part II — Cambridge International AS Level', 'Candidate No. MZ040/0062'),
      table([gradeHeader(), ...ASL.map(gradeRow), gradeFoot('Four syllabuses — AS Level stage average', '62.8%', '2.00')], GRADE_COLS),
      gap(110),

      resultPanel,
      gap(80),

      new Paragraph({
        alignment: AlignmentType.CENTER, spacing: { after: 60, line: 250 },
        indent: { left: 500, right: 500 },
        children: [
          run('Cumulative Grade Point Average: ', { size: 17, color: INK_SOFT }),
          run('Two point six four', { size: 17, bold: true }),
          run(' out of a maximum of ', { size: 17, color: INK_SOFT }),
          run('four point zero zero', { size: 17, bold: true }),
          run('. Aggregate percentage uniform mark across all eleven syllabuses: ', { size: 17, color: INK_SOFT }),
          run('69.55%', { size: 17, bold: true }),
          run('.', { size: 17, color: INK_SOFT }),
        ],
      }),
      new Paragraph({
        alignment: AlignmentType.CENTER, spacing: { after: 0, line: 250 },
        children: [
          run('Conversion applied: ', { size: 16, color: INK_FAINT }),
          run('A*/A = 4.00  ·  B = 3.00  ·  C = 2.00  ·  D = 1.00  ·  E and below = 0.00', { size: 16, bold: true, color: INK_SOFT }),
          run('. Full basis of calculation is set out in the annex overleaf.', { size: 16, color: INK_FAINT }),
        ],
      }),

      gap(60),
      signatureBlock([['Registrar', 'Name and signature'], ['Official seal', 'Affix seal here'], ['Principal', 'Name and signature']]),

      /* ===== PAGE 2 ===== */
      new Paragraph({ children: [new PageBreak()], spacing: { after: 0 } }),

      new Paragraph({
        spacing: { after: 60 },
        children: [
          run('Annex — Basis of Calculation', { font: SERIF, size: 30, bold: true, color: ACCENT, caps: true, track: 10 }),
        ],
      }),
      new Paragraph({
        spacing: { after: 160 },
        border: { bottom: line(RULE, 6) },
        alignment: AlignmentType.RIGHT,
        children: [run('Xavier Keylane Arianne Vasconcelos', { size: 16, bold: true, color: INK_SOFT }), run('  ·  Centre MZ040  ·  Ref. No. ______________', { size: 16, color: INK_FAINT })],
      }),

      new Paragraph({
        spacing: { before: 40, after: 160, line: 245 },
        shading: { type: ShadingType.CLEAR, fill: ACCENT_SOFT, color: 'auto' },
        border: { left: line(ACCENT, 18) },
        indent: { left: 160, right: 160 },
        children: [
          run('Result certified: ', { size: 18, bold: true, color: ACCENT }),
          run('Cumulative Grade Point Average of ', { size: 18, color: INK_SOFT }),
          run('2.64 on a 4.00 scale', { size: 18, bold: true, color: INK }),
          run(', derived from eleven Cambridge International syllabus grades awarded between June 2024 and November 2025.', { size: 18, color: INK_SOFT }),
        ],
      }),

      keyCalcBlock,
      gap(120),

      sectionHead('Method and scope'),
      methodPara('Method.', 'Each reported syllabus grade is converted to grade points using the key above. The cumulative grade point average is the arithmetic mean of the grade points of all eleven reported syllabuses, carried to two decimal places, with no intermediate rounding.'),
      methodPara('Unweighted CGPA — 2.64 / 4.00.', 'All eleven syllabuses carry equal weight (29.00 ÷ 11). This is the figure certified overleaf.'),
      methodPara('Credit-weighted alternative — 2.47 / 4.00.', 'Where an institution weights AS Level at two credits against one credit for IGCSE, the result is 37.00 ÷ 15.'),
      methodPara('Ten-point scale.', 'The unweighted result expressed on a 10.00 scale is 6.59 / 10.00; the credit-weighted result is 6.17 / 10.00.'),
      methodPara('Percentage uniform marks.', 'The PUM shown against each syllabus is the mark reported by Cambridge on the corresponding Statement of Results. It is a position-within-grade indicator, not a raw examination mark, and is reproduced here for reference only; it does not enter the grade point calculation. The aggregate PUM across the eleven syllabuses is 69.55%.'),
      methodPara('Scope.', 'Cambridge International does not itself issue grade point averages. This conversion is prepared by the issuing institution named overleaf, on the scale stated above. Receiving institutions that publish a different conversion scale should apply their own.'),
      gap(70),

      sectionHead('Source documents'),
      srcTable,
      gap(140),

      methodPara('Verification.', 'The grades certified in this document may be verified against the original Cambridge International certificates and statements of results listed above, and through the issuing institution’s Office of the Registrar quoting the reference number shown.'),

      gap(150),
      signatureBlock([['Registrar', 'Name and signature'], ['Date', 'Day / month / year']]),
    ],
  }],
});

Packer.toBuffer(doc).then((buf) => {
  fs.writeFileSync(process.argv[2] || 'out.docx', buf);
  console.log('written', process.argv[2], buf.length, 'bytes');
});
