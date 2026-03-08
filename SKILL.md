---
name: katzav-skill
description: >
  יצירת מסמכים משפטיים בעברית בפורמט DOCX עבור משרד אקוקא-קצב עורכי דין.
  תמיכה מלאה ב-RTL, עקוב אחר שינויים, הערות, נייר פירמה עם לוגו.

  טריגרים: "מסמך משפטי", "הסכם", "כתב הגנה", "כתב תביעה", "בקשה", "תצהיר",
  "מכתב התראה", "חוזה", "הסכם שירותים", "ייפוי כוח", "פרוטוקול", "החלטה",
  "צו", "פסק דין", "כתב טענות", "כתב בי דין", בקשה ליצור מסמך DOCX בעברית,
  "tracked changes בעברית", "הערות שוליים משפטיות", "עקוב אחר שינויים",
  "קצב סקיל", "katzav".
---

# קצב סקיל — מסמכים משפטיים | אקוקא-קצב משרד עו"ד

סקיל זה מרחיב את סקיל docx הבסיסי עם התמחות במסמכים משפטיים ישראליים,
מותאם ספציפית למשרד **אקוקא-קצב ושות' עורכי דין**.

**תמיד לקרוא קודם** את `/mnt/skills/public/docx/SKILL.md` — הסקיל הזה מניח שאתה מכיר את תהליך העבודה הבסיסי (docx-js, unpack/pack, tracked changes, comments).

---

## פרטי המשרד — FIRM CONFIGURATION

```javascript
const FIRM = {
  name: 'אקוקא-קצב, משרד עו"ד',
  nameEn: 'AKOKA-KATZAV, Law Office',

  attorneys: [
    { name: 'לירן אקוקא', license: '52786', nameEn: 'LIRAN AKOKA' },
    { name: 'אלירן קצב', license: '62333', nameEn: 'ELIRAN KATZAV' },
  ],

  address: 'רח\' סוקולוב 40, (בית הפירמידה), רמת השרון',
  addressEn: 'Sokolov 40, (Pyramid House), Ramat HaSharon',

  phone: '077-2088182',
  fax: '077-555-88-97',
  email: 'office@ak-law.co.il',
  website: 'www.ak-law.co.il',

  // שורת ייצוג בכתבי בי דין
  representationLine: 'ע"י ב"כ עוה"ד לירן אקוקא (52786) ו/או אלירן קצב (62333)',
  firmLine: 'אקוקא-קצב, משרד עו"ד',
  firmAddress: 'רח\' סוקולוב 40, (בית הפירמידה), רמת השרון',
  firmContact: 'טל\': 077-2088182 | פקס: 077-555-88-97 | office@ak-law.co.il',

  // נתיב ללוגו (מתוך assets של הסקיל)
  logoPath: 'assets/logo.png',

  // צבעי המשרד
  colors: {
    primary: '1B2A4A',   // כחול כהה
    secondary: 'C9A84C', // זהב
    gray: '555555',
    lightGray: '777777',
  }
};
```

---

## כלל מרכזי: מייצג תובע או נתבע?

**כשהמשתמש אומר שהוא מייצג תובע** — פרטי המשרד מופיעים תחת **התובע**:
```
התובע: [שם הלקוח], ת.ז. [מספר]
       ע"י ב"כ עוה"ד לירן אקוקא (52786) ו/או אלירן קצב (62333)
       אקוקא-קצב, משרד עו"ד
       רח' סוקולוב 40, (בית הפירמידה), רמת השרון
       טל': 077-2088182 | פקס: 077-555-88-97 | office@ak-law.co.il

                                   - נגד -

הנתבע: [שם הצד שכנגד], ת.ז./ח.פ. [מספר]
```

**כשהמשתמש אומר שהוא מייצג נתבע** — פרטי המשרד מופיעים תחת **הנתבע**:
```
התובע: [שם הצד שכנגד], ת.ז./ח.פ. [מספר]

                                   - נגד -

הנתבע: [שם הלקוח], ת.ז. [מספר]
        ע"י ב"כ עוה"ד לירן אקוקא (52786) ו/או אלירן קצב (62333)
        אקוקא-קצב, משרד עו"ד
        רח' סוקולוב 40, (בית הפירמידה), רמת השרון
        טל': 077-2088182 | פקס: 077-555-88-97 | office@ak-law.co.il
```

**אותו דבר חל על כל סוגי הצדדים:** מבקש/משיב, מערער/משיב, עותר/משיב — תמיד הצמד את פרטי המשרד לצד שאנחנו מייצגים.

---

## כללי RTL קריטיים

### הכלל המרכזי: START/END במקום LEFT/RIGHT

**במסמך עברי עם `bidirectional: true`, לעולם אל תשתמש ב-`AlignmentType.LEFT` או `AlignmentType.RIGHT` לפסקאות ומספור!**

| רוצה יישור ל... | לא להשתמש | להשתמש |
|-----------------|-------------|----------|
| **ימין** | `LEFT` או `RIGHT` | `AlignmentType.START` |
| **שמאל** | `LEFT` או `RIGHT` | `AlignmentType.END` |
| **מרכז** | — | `AlignmentType.CENTER` |
| **שני צדדים** | — | `AlignmentType.BOTH` |

### שלוש הגדרות RTL חובה

כל מסמך עברי חייב את **שלושת** ההגדרות הבאות בכל הרמות:

```javascript
// 1. ברמת ה-Section
sections: [{
  properties: {
    bidi: true  // חובה!
  }
}]

// 2. ברמת כל Paragraph
new Paragraph({
  bidirectional: true,  // חובה!
  alignment: AlignmentType.BOTH,
})

// 3. ברמת כל TextRun
new TextRun({
  text: "טקסט בעברית",
  rightToLeft: true,  // חובה!
  font: "David",
})
```

**חוסר באחת מהן = יישור שגוי או טקסט הפוך!**

---

## זיהוי סוג מסמך

**לפני יצירת מסמך, זהה את סוגו:**

| סוג מסמך | דוגמאות | Header בית משפט? | מבנה מיוחד |
|----------|---------|------------------|------------|
| **כתב טענות** | תביעה, הגנה, בקשה, ערעור, תצהיר, בר"ע | כן | טבלת Header עם בית משפט + מספר תיק |
| **מכתב התראה** | התראה, דרישה, מכתב עו"ד | לא | נייר פירמה + לוגו, "הנדון:", חתימה |
| **הסכם/חוזה** | הסכם שירותים, NDA, חוזה שכירות | לא | הואילים, צדדים, חתימות בשני טורים |
| **מסמך כללי** | חוות דעת, מזכר, סיכום | לא | נייר פירמה לפי הצורך |

---

## פונטים ומידות

### פונטים
| פונט | שימוש | size (half-points) |
|------|-------|-------------------|
| **David** | ברירת מחדל, גוף טקסט | 24 (12pt) |
| **Tahoma** | נייר פירמה (header/footer) | 20 (10pt) |

### מידות ושוליים
```
2.0 ס"מ = 1134 DXA (שוליים ברירת מחדל של המשרד)
2.5 ס"מ = 1417 DXA (חלופי)
A4 = 11906 × 16838 DXA
רוחב תוכן (שוליים 2.0 ס"מ) = 9638 DXA
```

---

## מספור סעיפים משפטיים

**alignment: AlignmentType.START — לא LEFT ולא RIGHT!**

```javascript
numbering: {
  config: [{
    reference: "legal-clauses",
    levels: [
      {
        level: 0,
        format: LevelFormat.DECIMAL,
        text: "%1.",
        alignment: AlignmentType.START,
        suffix: "tab",
        style: { paragraph: { indent: { left: 720, hanging: 360 } } }
      },
      {
        level: 1,
        format: LevelFormat.DECIMAL,
        text: "%1.%2",
        alignment: AlignmentType.START,
        suffix: "tab",
        style: { paragraph: { indent: { left: 1440, hanging: 500 } } }
      },
      {
        level: 2,
        format: LevelFormat.DECIMAL,
        text: "%1.%2.%3",
        alignment: AlignmentType.START,
        suffix: "tab",
        style: { paragraph: { indent: { left: 2160, hanging: 640 } } }
      }
    ]
  }]
}
```

---

## טבלאות RTL

**קריטי: `visuallyRightToLeft: true` — בלי זה העמודות הפוכות!**

```javascript
const CONTENT_WIDTH = 9638;  // A4 עם שוליים 2.0 ס"מ
const border = { style: BorderStyle.SINGLE, size: 1, color: "999999" };
const borders = { top: border, bottom: border, left: border, right: border };
const visibleBorder = { style: BorderStyle.SINGLE, size: 1, color: "999999" };
const visibleBorders = { top: visibleBorder, bottom: visibleBorder, left: visibleBorder, right: visibleBorder };

const rtlCell = (text, width, opts = {}) => new TableCell({
  borders: opts.noBorders ? noBorders : borders,
  width: { size: width, type: WidthType.DXA },
  margins: { top: 80, bottom: 80, left: 120, right: 120 },
  ...(opts.shading ? { shading: { fill: opts.shading, type: ShadingType.CLEAR } } : {}),
  children: [new Paragraph({
    bidirectional: true,
    alignment: opts.alignment || AlignmentType.CENTER,
    children: [new TextRun({
      text, font: "David", size: 24, rightToLeft: true, bold: opts.bold
    })]
  })]
});

new Table({
  visuallyRightToLeft: true,
  width: { size: CONTENT_WIDTH, type: WidthType.DXA },
  columnWidths: [4819, 2410, 2409],
  rows: [/* ... */]
})
```

---

## תבנית 1: כתב בי דין (תביעה / הגנה / בקשה)

זו התבנית העיקרית. כל ה-Header בנוי כ**טבלה אחת עם 3 עמודות וגבולות נראים (visibleBorders — אפור 999999)** — כמקובל במסמכים משפטיים ישראליים.

**מבנה:** עמודה ימנית (תוויות 2800), עמודה אמצעית (תוכן 5000), עמודה שמאלית (מספר תיק 1838).

**שים לב:** פרטי המשרד מוצמדים לצד שאנחנו מייצגים. עיר בית המשפט (`courtCity`) תמיד בשורה נפרדת.

```javascript
const { Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
        AlignmentType, LevelFormat, BorderStyle, WidthType, VerticalAlign, LineRuleType } = require('docx');

const PAGE_WIDTH = 11906;
const MARGINS = { top: 1134, right: 1134, bottom: 1134, left: 1134 };
const CONTENT_WIDTH = PAGE_WIDTH - MARGINS.left - MARGINS.right; // 9638

// === 3 עמודות עם visuallyRightToLeft: true ===
// columnWidths[0] = עמודה ימנית ויזואלית = תוויות (בית משפט, התובע:, הנתבע:)
// columnWidths[1] = עמודה אמצעית ויזואלית = תוכן (שמות, פרטים, ייצוג)
// columnWidths[2] = עמודה שמאלית ויזואלית = מספר תיק (ריקה ברוב השורות)
const COL_LABELS = 2800;   // ימין — תוויות
const COL_CONTENT = 5000;  // אמצע — תוכן
const COL_CASE = 1838;     // שמאל — מספר תיק
// Total: 2800 + 5000 + 1838 = 9638 ✓

const visibleBorder = { style: BorderStyle.SINGLE, size: 1, color: "999999" };
const visibleBorders = { top: visibleBorder, bottom: visibleBorder, left: visibleBorder, right: visibleBorder };

// גבולות נסתרים לכותרת כתב בי דין — נראים כ-gridlines בעורך Word אבל לא מודפסים
const noBorder = { style: BorderStyle.NONE, size: 0, color: "FFFFFF" };
const noBorders = { top: noBorder, bottom: noBorder, left: noBorder, right: noBorder };
// margins אחידים לתאי הכותרת
const headerCellMargins = { top: 40, bottom: 40, left: 80, right: 80 };

// === תוכן תא — צד מיוצג (כולל פרטי משרד) ===
function partyContentWithRepresentation(clientName, clientId, clientAddress) {
  return [
    new Paragraph({
      bidirectional: true, alignment: AlignmentType.START,
      spacing: { after: 40 },
      children: [
        new TextRun({ text: `${clientName}, ת.ז. ${clientId}`, font: "David", size: 24, rightToLeft: true, bold: true }),
      ]
    }),
    ...(clientAddress ? [new Paragraph({
      bidirectional: true, alignment: AlignmentType.START,
      spacing: { after: 40 },
      children: [new TextRun({ text: clientAddress, font: "David", size: 24, rightToLeft: true })]
    })] : []),
    new Paragraph({
      bidirectional: true, alignment: AlignmentType.START,
      spacing: { after: 40 },
      children: [new TextRun({
        text: 'ע"י ב"כ עוה"ד לירן אקוקא (52786) ו/או אלירן קצב (62333)',
        font: "David", size: 24, rightToLeft: true
      })]
    }),
    new Paragraph({
      bidirectional: true, alignment: AlignmentType.START,
      spacing: { after: 40 },
      children: [new TextRun({
        text: 'אקוקא-קצב, משרד עו"ד',
        font: "David", size: 24, rightToLeft: true
      })]
    }),
    new Paragraph({
      bidirectional: true, alignment: AlignmentType.START,
      spacing: { after: 40 },
      children: [new TextRun({
        text: "רח' סוקולוב 40, (בית הפירמידה), רמת השרון",
        font: "David", size: 24, rightToLeft: true
      })]
    }),
    new Paragraph({
      bidirectional: true, alignment: AlignmentType.START,
      spacing: { after: 40 },
      children: [new TextRun({
        text: "טל': 077-2088182 | פקס: 077-555-88-97",
        font: "David", size: 24, rightToLeft: true
      })]
    }),
    new Paragraph({
      bidirectional: true, alignment: AlignmentType.START,
      children: [new TextRun({
        text: 'דוא"ל: office@ak-law.co.il',
        font: "David", size: 24, rightToLeft: true
      })]
    }),
  ];
}

// === תוכן תא — צד שכנגד (עם/בלי ייצוג) ===
function partyContentOpposing(name, id, address, opposingRepresentation) {
  const paragraphs = [
    new Paragraph({
      bidirectional: true, alignment: AlignmentType.START,
      spacing: { after: 40 },
      children: [
        new TextRun({ text: `${name}, ת.ז./ח.פ. ${id}`, font: "David", size: 24, rightToLeft: true, bold: true }),
      ]
    }),
  ];
  if (address) {
    paragraphs.push(new Paragraph({
      bidirectional: true, alignment: AlignmentType.START,
      spacing: { after: 40 },
      children: [new TextRun({ text: address, font: "David", size: 24, rightToLeft: true })]
    }));
  }
  if (opposingRepresentation) {
    if (Array.isArray(opposingRepresentation)) {
      opposingRepresentation.forEach(line => {
        paragraphs.push(new Paragraph({
          bidirectional: true, alignment: AlignmentType.START,
          spacing: { after: 40 },
          children: [new TextRun({ text: line, font: "David", size: 24, rightToLeft: true })]
        }));
      });
    } else {
      paragraphs.push(new Paragraph({
        bidirectional: true, alignment: AlignmentType.START,
        spacing: { after: 40 },
        children: [new TextRun({ text: opposingRepresentation, font: "David", size: 24, rightToLeft: true })]
      }));
    }
  }
  return paragraphs;
}

// === טבלת Header כתב בי דין — טבלה אחת עם גבולות נסתרים (noBorders) ===
// הגבולות נראים כ-gridlines בעורך Word אבל לא מודפסים
// מבנה: 7 שורות × 3 עמודות (ימנית=תוויות, אמצעית=תוכן, שמאלית=מספר תיק)
// שורות 3, 5 ו-7 משתרעות על 3 עמודות (columnSpan: 3)
function courtDocumentHeader({
  courtName,         // "בבית המשפט השלום"
  courtCity,         // "בחיפה" — תמיד בשורה נפרדת מתחת לשם בית המשפט
  caseNumber,        // 'ת.א. ________/__/__'
  plaintiffContent,  // Paragraph[] מ-partyContentWithRepresentation או partyContentOpposing
  plaintiffLabel,    // "התובע" / "התובעים" / "המבקש"
  defendantContent,  // Paragraph[] מ-partyContentWithRepresentation או partyContentOpposing
  defendantLabel,    // "הנתבע" / "הנתבעת" / "הנתבעים" / "המשיב"
  documentTitle,     // "כתב תביעה" / "כתב הגנה" / "בקשה"
}) {
  return new Table({
    width: { size: CONTENT_WIDTH, type: WidthType.DXA },
    columnWidths: [COL_LABELS, COL_CONTENT, COL_CASE],
    visuallyRightToLeft: true,
    rows: [
      // === שורה 1: ת.חתימה (ימין) | ריק (אמצע) | ריק (שמאל) ===
      new TableRow({ children: [
        new TableCell({
          width: { size: COL_LABELS, type: WidthType.DXA }, borders: noBorders,
          margins: headerCellMargins,
          children: [new Paragraph({
            bidirectional: true, alignment: AlignmentType.START,
            spacing: { before: 0, after: 0 },
            children: [new TextRun({ text: "ת.חתימה :", font: "David", size: 24, rightToLeft: true })]
          })]
        }),
        new TableCell({
          width: { size: COL_CONTENT, type: WidthType.DXA }, borders: noBorders,
          margins: headerCellMargins,
          children: [new Paragraph({ bidirectional: true, spacing: { before: 0, after: 0 }, children: [] })]
        }),
        new TableCell({
          width: { size: COL_CASE, type: WidthType.DXA }, borders: noBorders,
          margins: headerCellMargins,
          children: [new Paragraph({ bidirectional: true, spacing: { before: 0, after: 0 }, children: [] })]
        }),
      ]}),
      // === שורה 2: בית משפט (ימין) | ריק (אמצע) | מספר תיק (שמאל) ===
      new TableRow({ children: [
        new TableCell({
          width: { size: COL_LABELS, type: WidthType.DXA }, borders: noBorders,
          margins: headerCellMargins,
          children: [
            new Paragraph({
              bidirectional: true, alignment: AlignmentType.START,
              spacing: { before: 0, after: 0 },
              children: [new TextRun({ text: courtName, bold: true, font: "David", size: 24, rightToLeft: true })]
            }),
            new Paragraph({
              bidirectional: true, alignment: AlignmentType.START,
              spacing: { before: 0, after: 0 },
              children: [new TextRun({ text: courtCity, bold: true, font: "David", size: 24, rightToLeft: true })]
            }),
          ]
        }),
        new TableCell({
          width: { size: COL_CONTENT, type: WidthType.DXA }, borders: noBorders,
          margins: headerCellMargins,
          children: [new Paragraph({ bidirectional: true, spacing: { before: 0, after: 0 }, children: [] })]
        }),
        new TableCell({
          width: { size: COL_CASE, type: WidthType.DXA }, borders: noBorders,
          margins: headerCellMargins,
          children: [new Paragraph({
            bidirectional: true, alignment: AlignmentType.START,
            spacing: { before: 0, after: 0 },
            children: [new TextRun({ text: caseNumber, bold: true, font: "David", size: 24, rightToLeft: true })]
          })]
        }),
      ]}),
      // === שורה מפרידה ריקה (colspan 3) ===
      new TableRow({ children: [
        new TableCell({
          width: { size: COL_LABELS, type: WidthType.DXA }, borders: noBorders,
          columnSpan: 3,
          children: [new Paragraph({ bidirectional: true, spacing: { before: 0, after: 0 }, children: [] })]
        }),
      ]}),
      // === שורה 4: תווית תובעים (ימין) | פרטי תובעים (אמצע) | ריק (שמאל) ===
      new TableRow({ children: [
        new TableCell({
          width: { size: COL_LABELS, type: WidthType.DXA }, borders: noBorders,
          margins: headerCellMargins,
          verticalAlign: VerticalAlign.TOP,
          children: [new Paragraph({
            bidirectional: true, alignment: AlignmentType.START,
            children: [new TextRun({ text: `${plaintiffLabel}:`, bold: true, font: "David", size: 24, rightToLeft: true, underline: {} })]
          })]
        }),
        new TableCell({
          width: { size: COL_CONTENT, type: WidthType.DXA }, borders: noBorders,
          margins: headerCellMargins,
          children: plaintiffContent
        }),
        new TableCell({
          width: { size: COL_CASE, type: WidthType.DXA }, borders: noBorders,
          margins: headerCellMargins,
          children: [new Paragraph({ bidirectional: true, children: [] })]
        }),
      ]}),
      // === שורה 5: "-נגד-" (colspan 3) ===
      new TableRow({ children: [
        new TableCell({
          width: { size: COL_LABELS, type: WidthType.DXA }, borders: noBorders,
          columnSpan: 3,
          children: [new Paragraph({
            bidirectional: true, alignment: AlignmentType.CENTER,
            spacing: { before: 200, after: 200 },
            children: [new TextRun({ text: "-נגד-", bold: true, font: "David", size: 24, rightToLeft: true })]
          })]
        }),
      ]}),
      // === שורה 6: תווית נתבע (ימין) | פרטי נתבע (אמצע) | ריק (שמאל) ===
      new TableRow({ children: [
        new TableCell({
          width: { size: COL_LABELS, type: WidthType.DXA }, borders: noBorders,
          margins: headerCellMargins,
          verticalAlign: VerticalAlign.TOP,
          children: [new Paragraph({
            bidirectional: true, alignment: AlignmentType.START,
            children: [new TextRun({ text: `${defendantLabel}:`, bold: true, font: "David", size: 24, rightToLeft: true, underline: {} })]
          })]
        }),
        new TableCell({
          width: { size: COL_CONTENT, type: WidthType.DXA }, borders: noBorders,
          margins: headerCellMargins,
          children: defendantContent
        }),
        new TableCell({
          width: { size: COL_CASE, type: WidthType.DXA }, borders: noBorders,
          margins: headerCellMargins,
          children: [new Paragraph({ bidirectional: true, children: [] })]
        }),
      ]}),
      // === שורה 7: כותרת המסמך (colspan 3) ===
      new TableRow({ children: [
        new TableCell({
          width: { size: COL_LABELS, type: WidthType.DXA }, borders: noBorders,
          columnSpan: 3,
          margins: headerCellMargins,
          children: [new Paragraph({
            bidirectional: true, alignment: AlignmentType.CENTER,
            spacing: { before: 300, after: 300 },
            children: [new TextRun({ text: documentTitle, bold: true, font: "David", size: 24, rightToLeft: true, underline: {} })]
          })]
        }),
      ]}),
    ]
  });
}

// === כותרת ראשית (גם לשימוש עצמאי — למשל בתצהיר) ===
function mainTitle(text) {
  return new Paragraph({
    bidirectional: true, alignment: AlignmentType.CENTER,
    spacing: { before: 300, after: 300 },
    children: [new TextRun({ text, bold: true, font: "David", size: 28, rightToLeft: true, underline: {} })]
  });
}

// === חתימת עו"ד ===
function lawyerSignature() {
  return [
    new Paragraph({
      bidirectional: true, alignment: AlignmentType.START,
      spacing: { before: 400 },
      children: [new TextRun({ text: "בכבוד רב,", font: "David", size: 24, rightToLeft: true })]
    }),
    new Paragraph({ bidirectional: true, spacing: { before: 200 }, children: [] }),
    new Paragraph({
      bidirectional: true, alignment: AlignmentType.START,
      children: [new TextRun({ text: "________________", font: "David", size: 24, rightToLeft: true })]
    }),
    new Paragraph({
      bidirectional: true, alignment: AlignmentType.START,
      children: [new TextRun({
        text: 'לירן אקוקא, עו"ד / אלירן קצב, עו"ד',
        bold: true, font: "David", size: 24, rightToLeft: true
      })]
    }),
    new Paragraph({
      bidirectional: true, alignment: AlignmentType.START,
      children: [new TextRun({
        text: 'אקוקא-קצב, משרד עו"ד',
        font: "David", size: 24, rightToLeft: true
      })]
    }),
  ];
}

// === דוגמה: כתב תביעה (מייצגים את התובע) ===
const doc = new Document({
  numbering: {
    config: [{
      reference: "legal-clauses",
      levels: [{
        level: 0, format: LevelFormat.DECIMAL, text: "%1.",
        alignment: AlignmentType.START, suffix: "tab",
        style: { paragraph: { indent: { left: 360, hanging: 360 } } }
      }]
    }]
  },
  sections: [{
    properties: {
      page: { size: { width: PAGE_WIDTH, height: 16838 }, margin: MARGINS },
      bidi: true
    },
    children: [
      courtDocumentHeader({
        courtName: "בבית המשפט השלום",
        courtCity: "בפתח תקווה",
        caseNumber: 'ת"א _____-__-__',
        plaintiffContent: partyContentWithRepresentation("[שם הלקוח]", "[מספר ת.ז.]", "[כתובת]"),
        plaintiffLabel: "התובע",
        defendantContent: partyContentOpposing("[שם הנתבע]", "[מספר]", "[כתובת]"),
        defendantLabel: "הנתבע",
        documentTitle: "כתב תביעה",
      }),
      // ... סעיפים ...
      ...lawyerSignature(),
    ]
  }]
});

// === דוגמה: כתב הגנה (מייצגים את הנתבע) ===
// children: [
//   courtDocumentHeader({
//     courtName: "בבית המשפט השלום",
//     courtCity: "בחיפה",
//     caseNumber: 'ת"א _____-__-__',
//     plaintiffContent: partyContentOpposing("[שם התובע]", "[מספר]", "[כתובת]", [
//       'ע"י ב"כ עו"ד [שם עו"ד] ([מ.ר.])',
//       'משרד [שם המשרד], עורכי דין',
//       '[כתובת המשרד]',
//       'טל: [טלפון] | פקס: [פקס]',
//       'דוא"ל: [אימייל]',
//     ]),
//     plaintiffLabel: "התובע",
//     defendantContent: partyContentWithRepresentation("[שם הלקוח]", "[מספר ת.ז.]", "[כתובת]"),
//     defendantLabel: "הנתבע",
//     documentTitle: "כתב הגנה",
//   }),
//   ...
// ]
```

---

## תבנית 2: מכתב התראה (עם נייר פירמה)

### שיטה מועדפת: unpack/repack מתבנית נייר מכתבים

**תמיד להשתמש בתבנית `assets/letterhead-template.docx` ליצירת מכתבים!**

התבנית כוללת נייר מכתבים מדויק עם:
- Header: לוגו המשרד (שמאל) + שמות עו"ד בעברית ואנגלית (ימין), פונט minorBidi 12pt, צבע 1B2A4A
- קו מפריד כחול כהה (1B2A4A, size 6)
- Footer: קו עליון כחול, כתובת בעברית ואנגלית, פרטי קשר (Tahoma 9pt, צבע 1F497D)
- שוליים: top 850, right/bottom/left 1134, header 283, footer 283
- גוף: פונט David 12pt (size 24), RTL

```bash
# שלב 1: unpack התבנית
python /mnt/skills/public/docx/scripts/unpack.py assets/letterhead-template.docx unpacked/

# שלב 2: ערוך word/document.xml — הוסף פסקאות תוכן לפני <w:sectPr>
# שלב 3: repack
python /mnt/skills/public/docx/scripts/pack.py unpacked/ output.docx --original assets/letterhead-template.docx
```

### פסקאות תוכן להכנסה (XML — לפני `<w:sectPr>`)

**חשוב:** החלף את כל הפסקאות הריקות הקיימות (שורות `<w:p>` הריקות שאחרי הטבלה וקו המפריד) בפסקאות התוכן הבאות:

```xml
<!-- תאריך -->
<w:p><w:pPr><w:bidi/><w:spacing w:after="0"/><w:jc w:val="right"/></w:pPr>
<w:r><w:rPr><w:rFonts w:ascii="David" w:hAnsi="David" w:cs="David"/>
<w:sz w:val="24"/><w:szCs w:val="24"/><w:rtl/></w:rPr>
<w:t>[תאריך]</w:t></w:r></w:p>

<!-- סימוכין -->
<w:p><w:pPr><w:bidi/><w:spacing w:after="100"/><w:jc w:val="right"/></w:pPr>
<w:r><w:rPr><w:rFonts w:ascii="David" w:hAnsi="David" w:cs="David"/>
<w:sz w:val="24"/><w:szCs w:val="24"/><w:rtl/></w:rPr>
<w:t>סימוכין: [מספר תיק]</w:t></w:r></w:p>

<!-- נמען -->
<w:p><w:pPr><w:bidi/><w:spacing w:before="200" w:after="40"/><w:jc w:val="right"/></w:pPr>
<w:r><w:rPr><w:rFonts w:ascii="David" w:hAnsi="David" w:cs="David"/>
<w:sz w:val="24"/><w:szCs w:val="24"/><w:rtl/></w:rPr>
<w:t>לכבוד</w:t></w:r></w:p>

<w:p><w:pPr><w:bidi/><w:jc w:val="right"/></w:pPr>
<w:r><w:rPr><w:rFonts w:ascii="David" w:hAnsi="David" w:cs="David"/>
<w:b/><w:bCs/><w:sz w:val="24"/><w:szCs w:val="24"/><w:rtl/></w:rPr>
<w:t>[שם הנמען]</w:t></w:r></w:p>

<!-- סיווג -->
<w:p><w:pPr><w:bidi/><w:spacing w:before="200"/><w:jc w:val="right"/></w:pPr>
<w:r><w:rPr><w:rFonts w:ascii="David" w:hAnsi="David" w:cs="David"/>
<w:b/><w:bCs/><w:sz w:val="24"/><w:szCs w:val="24"/><w:rtl/></w:rPr>
<w:t>מכתב זה נשלח מבלי לפגוע בזכויות מרשי/תי</w:t></w:r></w:p>

<!-- הנדון -->
<w:p><w:pPr><w:bidi/><w:spacing w:before="200" w:after="200"/><w:jc w:val="center"/></w:pPr>
<w:r><w:rPr><w:rFonts w:ascii="David" w:hAnsi="David" w:cs="David"/>
<w:b/><w:bCs/><w:sz w:val="24"/><w:szCs w:val="24"/><w:rtl/></w:rPr>
<w:t xml:space="preserve">הנדון: </w:t></w:r>
<w:r><w:rPr><w:rFonts w:ascii="David" w:hAnsi="David" w:cs="David"/>
<w:b/><w:bCs/><w:u w:val="single"/><w:sz w:val="24"/><w:szCs w:val="24"/><w:rtl/></w:rPr>
<w:t>[נושא המכתב]</w:t></w:r></w:p>

<!-- גוף המכתב — סעיפים ממוספרים או פסקאות -->
<!-- ... הוסף פסקאות תוכן כאן ... -->

<!-- חתימה -->
<w:p><w:pPr><w:bidi/><w:spacing w:before="400"/><w:jc w:val="right"/></w:pPr>
<w:r><w:rPr><w:rFonts w:ascii="David" w:hAnsi="David" w:cs="David"/>
<w:sz w:val="24"/><w:szCs w:val="24"/><w:rtl/></w:rPr>
<w:t>בכבוד רב,</w:t></w:r></w:p>

<w:p><w:pPr><w:bidi/><w:spacing w:before="200"/></w:pPr></w:p>

<w:p><w:pPr><w:bidi/><w:jc w:val="right"/></w:pPr>
<w:r><w:rPr><w:rFonts w:ascii="David" w:hAnsi="David" w:cs="David"/>
<w:sz w:val="24"/><w:szCs w:val="24"/><w:rtl/></w:rPr>
<w:t>________________</w:t></w:r></w:p>

<w:p><w:pPr><w:bidi/><w:jc w:val="right"/></w:pPr>
<w:r><w:rPr><w:rFonts w:ascii="David" w:hAnsi="David" w:cs="David"/>
<w:b/><w:bCs/><w:sz w:val="24"/><w:szCs w:val="24"/><w:rtl/></w:rPr>
<w:t>לירן אקוקא, עו"ד / אלירן קצב, עו"ד</w:t></w:r></w:p>

<w:p><w:pPr><w:bidi/><w:jc w:val="right"/></w:pPr>
<w:r><w:rPr><w:rFonts w:ascii="David" w:hAnsi="David" w:cs="David"/>
<w:sz w:val="24"/><w:szCs w:val="24"/><w:rtl/></w:rPr>
<w:t>אקוקא-קצב, משרד עו"ד</w:t></w:r></w:p>
```

### שיטה חלופית: docx-js מאפס (אם אין גישה ל-unpack/pack)

```javascript
// אם אין אפשרות ל-unpack/repack, השתמש בקוד:
const doc = new Document({
  sections: [{
    properties: {
      page: { size: { width: 11906, height: 16838 },
              margin: { top: 850, right: 1134, bottom: 1134, left: 1134, header: 283, footer: 283 } },
      bidi: true
    },
    // ... Header ו-Footer כמו בתבנית המקורית ...
    children: [
      // תאריך, סימוכין, נמען, סיווג, הנדון, גוף, חתימה
    ]
  }]
});
```

---

## תבנית 3: הסכם/חוזה

```javascript
function contractTitle(text) {
  return new Paragraph({
    bidirectional: true, alignment: AlignmentType.CENTER,
    spacing: { after: 300 },
    children: [new TextRun({ text, bold: true, font: "David", size: 32, rightToLeft: true })]
  });
}

function partyClause(label, name, id, address, alias) {
  return new Paragraph({
    bidirectional: true, alignment: AlignmentType.BOTH,
    spacing: { after: 120 },
    children: [
      new TextRun({ text: `${label}: `, bold: true, font: "David", size: 24, rightToLeft: true }),
      new TextRun({ text: `${name}, ח.פ./ת.ז. ${id}, מ${address} (להלן: "`, font: "David", size: 24, rightToLeft: true }),
      new TextRun({ text: alias, bold: true, font: "David", size: 24, rightToLeft: true }),
      new TextRun({ text: '")', font: "David", size: 24, rightToLeft: true }),
    ]
  });
}

function signatureTable() {
  return new Table({
    width: { size: CONTENT_WIDTH, type: WidthType.DXA },
    columnWidths: [CONTENT_WIDTH / 2, CONTENT_WIDTH / 2],
    visuallyRightToLeft: true,
    rows: [new TableRow({ children: [
      new TableCell({ borders: visibleBorders, children: [
        new Paragraph({ bidirectional: true, alignment: AlignmentType.CENTER,
          children: [new TextRun({ text: "_________________", font: "David", size: 24, rightToLeft: true })] }),
        new Paragraph({ bidirectional: true, alignment: AlignmentType.CENTER,
          children: [new TextRun({ text: "צד א'", font: "David", size: 24, rightToLeft: true })] })
      ]}),
      new TableCell({ borders: visibleBorders, children: [
        new Paragraph({ bidirectional: true, alignment: AlignmentType.CENTER,
          children: [new TextRun({ text: "_________________", font: "David", size: 24, rightToLeft: true })] }),
        new Paragraph({ bidirectional: true, alignment: AlignmentType.CENTER,
          children: [new TextRun({ text: "צד ב'", font: "David", size: 24, rightToLeft: true })] })
      ]})
    ]})]
  });
}
```

---

## תבנית 4: תצהיר

```javascript
function affidavitOpening(name, idNumber) {
  return [
    mainTitle("תצהיר"),
    new Paragraph({
      bidirectional: true, alignment: AlignmentType.BOTH,
      spacing: { after: 200, line: 360, lineRule: LineRuleType.AUTO },
      children: [new TextRun({
        text: `אני, ${name}, ת.ז. ${idNumber}, לאחר שהוזהרתי כי עלי לומר את האמת וכי אהיה צפוי/ה לעונשים הקבועים בחוק אם לא אעשה כן, מצהיר/ה בזה כדלקמן:`,
        font: "David", size: 24, rightToLeft: true
      })]
    }),
  ];
}

function affidavitVerification(declarantName, idNumber, date) {
  return [
    new Paragraph({
      bidirectional: true, alignment: AlignmentType.START,
      spacing: { before: 300, after: 60 },
      children: [new TextRun({ text: "________________", font: "David", size: 24, rightToLeft: true })]
    }),
    new Paragraph({
      bidirectional: true, alignment: AlignmentType.START,
      spacing: { after: 300 },
      children: [new TextRun({ text: declarantName, bold: true, font: "David", size: 24, rightToLeft: true })]
    }),
    new Paragraph({
      bidirectional: true, alignment: AlignmentType.CENTER,
      spacing: { before: 200, after: 200 },
      children: [new TextRun({ text: "אישור עורך דין", bold: true, font: "David", size: 26, rightToLeft: true, underline: {} })]
    }),
    new Paragraph({
      bidirectional: true, alignment: AlignmentType.BOTH,
      spacing: { line: 360, lineRule: LineRuleType.AUTO },
      children: [new TextRun({
        text: `אני, עו"ד לירן אקוקא / אלירן קצב, מאשר/ת בזה כי ביום ${date} התייצב/ה בפני ${declarantName}, שזיהה/תה עצמו/ה באמצעות ת.ז. ${idNumber}, ולאחר שהזהרתיו/ה כי עליו/ה להצהיר אמת וכי יהיה/תהיה צפוי/ה לעונשים הקבועים בחוק אם לא יעשה/תעשה כן, אישר/ה את נכונות הצהרתו/ה וחתם/ה עליה בפני.`,
        font: "David", size: 24, rightToLeft: true
      })]
    }),
    new Paragraph({
      bidirectional: true, alignment: AlignmentType.START,
      spacing: { before: 200 },
      children: [new TextRun({ text: "________________", font: "David", size: 24, rightToLeft: true })]
    }),
    new Paragraph({
      bidirectional: true, alignment: AlignmentType.START,
      children: [new TextRun({
        text: 'עו"ד לירן אקוקא / אלירן קצב',
        bold: true, font: "David", size: 24, rightToLeft: true
      })]
    }),
  ];
}
```

---

## Tracked Changes — עקוב אחר שינויים

### שם מחבר
```xml
<w:del w:id="10" w:author="עו&quot;ד אקוקא" w:date="2026-03-07T09:00:00Z">
```

### RTL PROPS — בלוק rPr מלא
```xml
<w:rPr>
  <w:rFonts w:ascii="David" w:cs="David" w:eastAsia="David" w:hAnsi="David"/>
  <w:sz w:val="24"/>
  <w:szCs w:val="24"/>
  <w:rtl/>
</w:rPr>
```

### שינוי ערך
```xml
<w:r><w:rPr>...RTL PROPS...</w:rPr>
  <w:t xml:space="preserve">שכר הטרחה יעמוד על סך של </w:t></w:r>
<w:del w:id="10" w:author="עו&quot;ד אקוקא" w:date="...">
  <w:r><w:rPr>...RTL PROPS...</w:rPr><w:delText>750</w:delText></w:r>
</w:del>
<w:ins w:id="11" w:author="עו&quot;ד אקוקא" w:date="...">
  <w:r><w:rPr>...RTL PROPS...</w:rPr><w:t>850</w:t></w:r>
</w:ins>
<w:r><w:rPr>...RTL PROPS...</w:rPr>
  <w:t xml:space="preserve"> ש״ח לשעת עבודה</w:t></w:r>
```

---

## הערות שוליים (Footnotes)

```javascript
const { FootnoteReferenceRun } = require('docx');

const doc = new Document({
  footnotes: {
    1: { children: [new Paragraph({
      bidirectional: true, alignment: AlignmentType.START,
      children: [new TextRun({
        text: "חוק החוזים (חלק כללי), התשל״ג-1973, סעיף 12.",
        font: "David", size: 20, rightToLeft: true
      })]
    })] },
  },
  // ...
});

// הפניה בגוף הטקסט:
new Paragraph({
  bidirectional: true, alignment: AlignmentType.BOTH,
  children: [
    new TextRun({ text: "חובת תום הלב", font: "David", size: 24, rightToLeft: true }),
    new FootnoteReferenceRun(1),
    new TextRun({ text: " חלה על כל שלבי המשא ומתן.", font: "David", size: 24, rightToLeft: true }),
  ]
})
```

---

## מרווח שורות

```javascript
const { LineRuleType } = require('docx');

// LineRuleType.AUTO — הערך הוא ב-1/240 שורה
spacing: { line: 240, lineRule: LineRuleType.AUTO }  // 1.0
spacing: { line: 276, lineRule: LineRuleType.AUTO }  // 1.15 — ברירת מחדל
spacing: { line: 360, lineRule: LineRuleType.AUTO }  // 1.5 — נדרש בבתי משפט
spacing: { line: 480, lineRule: LineRuleType.AUTO }  // 2.0
```

---

## עריכת DOCX קיים

```bash
python /mnt/skills/public/docx/scripts/unpack.py input.docx unpacked/
# עריכת word/document.xml
python /mnt/skills/public/docx/scripts/pack.py unpacked/ output.docx --original input.docx
```

**כלל קריטי:** פסקאות חדשות חייבות להיכנס *לפני* `<w:sectPr>` האחרון.

---

## Quick Reference

### יישור
| רוצה | השתמש ב... |
|------|-----------|
| ימין | `AlignmentType.START` |
| שמאל | `AlignmentType.END` |
| מרכז | `AlignmentType.CENTER` |
| שני צדדים | `AlignmentType.BOTH` |

### גדלי טקסט
| שימוש | size | נקודות |
|-------|------|--------|
| גוף טקסט | 24 | 12pt |
| כותרת משנה | 26 | 13pt |
| כותרת ראשית | 28-32 | 14-16pt |
| הערות שוליים | 20 | 10pt |
| Header/Footer | 15-20 | 7.5-10pt |

### Checklist — הגדרות חובה
```
[ ] Section: bidi: true
[ ] Paragraph: bidirectional: true
[ ] TextRun: rightToLeft: true
[ ] Numbering: alignment: AlignmentType.START
[ ] Table: visuallyRightToLeft: true
[ ] פרטי משרד תחת הצד הנכון (תובע/נתבע)
```

---

## Troubleshooting

| בעיה | פתרון |
|------|-------|
| מספור הפוך (.1) | `alignment: AlignmentType.START` |
| טקסט שמאל-ימין | שלושת הגדרות RTL חסרות |
| עמודות טבלה הפוכות | `visuallyRightToLeft: true` |
| DOCX לא נפתח | לא להוסיף פסקה אחרי `sectPr` |
| הערות שוליים לא RTL | תקן ב-footnotes.xml אחרי unpack |

---

## קבצי עזר

- **`references/document-types.md`** — מבנים מפורטים ל-9 סוגי מסמכים משפטיים
- **`scripts/create-legal-doc.js`** — סקריפט עם פרטי המשרד מוטמעים
- **`assets/logo.png`** — לוגו המשרד לנייר פירמה

## Dependencies

```bash
npm install docx
```
