#!/usr/bin/env node
/**
 * create-legal-doc.js — תבנית מסמך משפטי | אקוקא-קצב משרד עו"ד
 *
 * שימוש: העתק לתיקיית העבודה, ערוך את CONTENT, הרץ עם node.
 *
 * כולל נייר פירמה מובנה עם פרטי המשרד.
 * RTL מלא ב-5 רמות: Document, Section, Paragraph, Run, Numbering.
 */

const fs = require('fs');
const path = require('path');
const { Document, Packer, Paragraph, TextRun, Header, Footer,
        AlignmentType, HeadingLevel, PageNumber, LevelFormat,
        Table, TableRow, TableCell, WidthType, BorderStyle,
        LineRuleType, ShadingType } = require('docx');

// ═══════════════════════════════════════════════
// FIRM — פרטי המשרד
// ═══════════════════════════════════════════════
const FIRM = {
  name: 'אקוקא-קצב, משרד עו"ד',
  attorneys: [
    { name: 'לירן אקוקא', license: '52786', nameEn: 'LIRAN AKOKA' },
    { name: 'אלירן קצב', license: '62333', nameEn: 'ELIRAN KATZAV' },
  ],
  address: "רח' סוקולוב 40, (בית הפירמידה), רמת השרון",
  addressEn: 'Sokolov 40, (Pyramid House), Ramat HaSharon',
  phone: '077-2088182',
  fax: '077-555-88-97',
  email: 'office@ak-law.co.il',
  representationLine: 'ע"י ב"כ עוה"ד לירן אקוקא (52786) ו/או אלירן קצב (62333)',
};

// ═══════════════════════════════════════════════
// CONFIGURATION
// ═══════════════════════════════════════════════
const FONT = "David";
const FONT_SIZE = 24;           // 12pt
const HEADING1_SIZE = 32;       // 16pt
const HEADING2_SIZE = 28;       // 14pt
const MARGIN_DXA = 1134;        // 2.0 ס"מ
const CONTENT_WIDTH = 11906 - MARGIN_DXA * 2; // 9638
const OUTPUT_FILE = "legal-document.docx";

const noBorder = { style: BorderStyle.NONE, size: 0, color: "FFFFFF" };
const noBorders = { top: noBorder, bottom: noBorder, left: noBorder, right: noBorder };

// ═══════════════════════════════════════════════
// HELPERS
// ═══════════════════════════════════════════════

const rtlRun = (text, opts = {}) => new TextRun({
  text, ...opts,
  font: opts.font || FONT,
  size: opts.size || FONT_SIZE,
  rightToLeft: true,
});

const rtlPara = (children, opts = {}) => new Paragraph({
  bidirectional: true,
  alignment: opts.alignment || AlignmentType.BOTH,
  spacing: opts.spacing,
  children: Array.isArray(children) ? children : [children],
  ...(opts.heading ? { heading: opts.heading } : {}),
  ...(opts.numbering ? { numbering: opts.numbering } : {}),
  ...(opts.indent ? { indent: opts.indent } : {}),
  ...(opts.border ? { border: opts.border } : {}),
});

const heading1 = (text) => rtlPara(
  rtlRun(text, { bold: true, size: HEADING1_SIZE, underline: {} }),
  { heading: HeadingLevel.HEADING_1, alignment: AlignmentType.CENTER }
);

const heading2 = (text) => rtlPara(
  rtlRun(text, { bold: true, size: HEADING2_SIZE }),
  { heading: HeadingLevel.HEADING_2, alignment: AlignmentType.START }
);

// === פרטי ייצוג (תחת הצד שמייצגים) ===
function representationBlock(role, clientName, clientId) {
  return [
    rtlPara([
      rtlRun(`${role}: `, { bold: true }),
      rtlRun(`${clientName}, ת.ז. ${clientId}`),
    ], { spacing: { after: 40 } }),
    rtlPara(rtlRun(FIRM.representationLine), {
      spacing: { after: 40 }, indent: { start: 720 }
    }),
    rtlPara(rtlRun(FIRM.name), {
      spacing: { after: 40 }, indent: { start: 720 }
    }),
    rtlPara(rtlRun(FIRM.address), {
      spacing: { after: 40 }, indent: { start: 720 }
    }),
    rtlPara(rtlRun(`טל': ${FIRM.phone} | פקס: ${FIRM.fax} | ${FIRM.email}`), {
      spacing: { after: 120 }, indent: { start: 720 }
    }),
  ];
}

// === צד שכנגד ===
function opposingParty(role, name, id) {
  return rtlPara([
    rtlRun(`${role}: `, { bold: true }),
    rtlRun(`${name}, ת.ז./ח.פ. ${id}`),
  ], { spacing: { after: 120 } });
}

// === "- נגד -" ===
function vsLine() {
  return rtlPara(rtlRun("- נגד -", { bold: true }), {
    alignment: AlignmentType.CENTER,
    spacing: { before: 200, after: 200 }
  });
}

// === חתימת עו"ד ===
function lawyerSignature() {
  return [
    rtlPara(rtlRun("בכבוד רב,"), { spacing: { before: 400 } }),
    rtlPara(rtlRun(""), { spacing: { before: 200 } }),
    rtlPara(rtlRun("________________")),
    rtlPara(rtlRun('לירן אקוקא, עו"ד / אלירן קצב, עו"ד', { bold: true })),
    rtlPara(rtlRun(FIRM.name)),
  ];
}

// === Header בית משפט ===
function courtHeader(courtName, caseNumber) {
  return new Table({
    width: { size: CONTENT_WIDTH, type: WidthType.DXA },
    columnWidths: [CONTENT_WIDTH / 2, CONTENT_WIDTH / 2],
    visuallyRightToLeft: true,
    rows: [new TableRow({ children: [
      new TableCell({
        width: { size: CONTENT_WIDTH / 2, type: WidthType.DXA }, borders: noBorders,
        children: [rtlPara(rtlRun(courtName, { bold: true, size: 26 }),
          { alignment: AlignmentType.START })]
      }),
      new TableCell({
        width: { size: CONTENT_WIDTH / 2, type: WidthType.DXA }, borders: noBorders,
        children: [rtlPara(rtlRun(caseNumber, { bold: true, size: 26 }),
          { alignment: AlignmentType.END })]
      })
    ]})]
  });
}

// === Footer נייר פירמה ===
function firmFooter() {
  return new Footer({
    children: [
      new Paragraph({ children: [] }),
      new Table({
        width: { size: CONTENT_WIDTH, type: WidthType.DXA },
        rows: [new TableRow({ children: [
          new TableCell({
            width: { size: CONTENT_WIDTH, type: WidthType.DXA }, borders: noBorders,
            shading: { fill: "F2F2F2", type: ShadingType.CLEAR },
            margins: { top: 100, bottom: 80, left: 200, right: 200 },
            children: [
              new Paragraph({
                alignment: AlignmentType.CENTER, spacing: { after: 20 },
                children: [new TextRun({
                  text: FIRM.address,
                  font: "Tahoma", size: 15, color: "555555", rightToLeft: true
                })]
              }),
              new Paragraph({
                alignment: AlignmentType.CENTER, spacing: { after: 20 },
                children: [new TextRun({
                  text: FIRM.addressEn,
                  font: "Tahoma", size: 14, color: "777777"
                })]
              }),
              new Paragraph({
                alignment: AlignmentType.CENTER,
                children: [
                  new TextRun({ text: FIRM.email, font: "Tahoma", size: 15, color: "1B2A4A" }),
                  new TextRun({ text: "  |  ", font: "Tahoma", size: 15, color: "777777" }),
                  new TextRun({ text: "פקס", font: "Tahoma", size: 15, color: "555555", rightToLeft: true }),
                  new TextRun({ text: `: ${FIRM.fax}  `, font: "Tahoma", size: 15, color: "555555" }),
                  new TextRun({ text: "|  ", font: "Tahoma", size: 15, color: "777777" }),
                  new TextRun({ text: "טל", font: "Tahoma", size: 15, color: "555555", rightToLeft: true }),
                  new TextRun({ text: `: ${FIRM.phone}`, font: "Tahoma", size: 15, color: "555555" }),
                ]
              }),
            ]
          })
        ]})]
      })
    ]
  });
}

// ═══════════════════════════════════════════════
// CONTENT — דוגמה: כתב תביעה (מייצגים את התובע)
// ═══════════════════════════════════════════════

const CONTENT = [
  courtHeader("בבית המשפט השלום ברמת השרון", 'ת"א _____-__-__'),

  // אנחנו מייצגים את התובע
  ...representationBlock("התובע", "[שם הלקוח]", "[מספר ת.ז.]"),
  vsLine(),
  opposingParty("הנתבע", "[שם הנתבע]", "[מספר]"),

  heading1("כתב תביעה"),

  heading2("מבוא"),
  rtlPara(rtlRun("[תוכן המבוא]"), { spacing: { after: 120 } }),

  heading2("העובדות"),
  rtlPara(rtlRun("1. [עובדה ראשונה]"), { spacing: { after: 120 } }),
  rtlPara(rtlRun("2. [עובדה שנייה]"), { spacing: { after: 120 } }),

  heading2("הסעדים המבוקשים"),
  rtlPara(rtlRun("[פירוט הסעדים]"), { spacing: { after: 120 } }),

  ...lawyerSignature(),
];

// ═══════════════════════════════════════════════
// DOCUMENT GENERATION
// ═══════════════════════════════════════════════

const doc = new Document({
  styles: {
    default: {
      document: {
        run: { font: FONT, size: FONT_SIZE, rightToLeft: true },
        paragraph: { bidirectional: true, alignment: AlignmentType.BOTH }
      }
    },
    paragraphStyles: [
      {
        id: "Heading1", name: "Heading 1", basedOn: "Normal", next: "Normal",
        quickFormat: true,
        run: { size: HEADING1_SIZE, bold: true, font: FONT, rightToLeft: true },
        paragraph: { spacing: { before: 240, after: 240 }, outlineLevel: 0,
                     bidirectional: true, alignment: AlignmentType.CENTER }
      },
      {
        id: "Heading2", name: "Heading 2", basedOn: "Normal", next: "Normal",
        quickFormat: true,
        run: { size: HEADING2_SIZE, bold: true, font: FONT, rightToLeft: true },
        paragraph: { spacing: { before: 200, after: 200 }, outlineLevel: 1,
                     bidirectional: true, alignment: AlignmentType.START }
      },
    ]
  },
  numbering: {
    config: [{
      reference: "legal-clauses",
      levels: [
        {
          level: 0, format: LevelFormat.DECIMAL, text: "%1.",
          alignment: AlignmentType.START, suffix: "tab",
          style: { paragraph: { indent: { left: 720, hanging: 360 } } }
        },
        {
          level: 1, format: LevelFormat.DECIMAL, text: "%1.%2",
          alignment: AlignmentType.START, suffix: "tab",
          style: { paragraph: { indent: { left: 1440, hanging: 500 } } }
        },
      ]
    }]
  },
  sections: [{
    properties: {
      page: {
        size: { width: 11906, height: 16838 },
        margin: { top: MARGIN_DXA, right: MARGIN_DXA, bottom: MARGIN_DXA, left: MARGIN_DXA }
      },
      bidi: true,
    },
    footers: { default: firmFooter() },
    children: CONTENT
  }]
});

Packer.toBuffer(doc).then(buffer => {
  fs.writeFileSync(OUTPUT_FILE, buffer);
  console.log(`✅ ${OUTPUT_FILE} created successfully`);
  console.log(`   Firm: ${FIRM.name}`);
  console.log(`   Font: ${FONT} ${FONT_SIZE/2}pt`);
  console.log(`   Size: ${(buffer.length / 1024).toFixed(1)} KB`);
  console.log(`   RTL: bidi + bidirectional + rightToLeft ✓`);
  console.log(`   Alignment: START/END (not LEFT/RIGHT) ✓`);
});
