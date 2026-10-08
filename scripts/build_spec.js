const {
  Document, Packer, Paragraph, TextRun, HeadingLevel, AlignmentType,
  Table, TableRow, TableCell, WidthType, BorderStyle, ShadingType,
  TableOfContents, PageBreak, Header, Footer, PageNumber, LevelFormat,
  ExternalHyperlink
} = require('docx');
const fs = require('fs');

// ---------- Style constants ----------
const FONT = 'Arial';
const NAVY = '1F3864';
const BLUE = '2E5496';
const ACCENT = 'A6611A';   // Roman-ish gold/ochre
const LIGHT = 'EAF0F7';
const GREY = 'F2F2F2';
const HEADSHADE = '1F3864';

// ---------- Helpers ----------
function p(text, opts = {}) {
  const runs = Array.isArray(text) ? text : [new TextRun({ text, rightToLeft: true, font: FONT, size: opts.size || 22, bold: opts.bold, italics: opts.italics, color: opts.color })];
  return new Paragraph({
    bidirectional: true,
    alignment: opts.align || AlignmentType.RIGHT,
    spacing: { after: opts.after != null ? opts.after : 120, line: 300, before: opts.before || 0 },
    children: runs,
  });
}

function h1(text) {
  return new Paragraph({
    heading: HeadingLevel.HEADING_1,
    bidirectional: true,
    alignment: AlignmentType.RIGHT,
    spacing: { before: 300, after: 160 },
    children: [new TextRun({ text, rightToLeft: true, font: FONT, size: 32, bold: true, color: NAVY })],
  });
}
function h2(text) {
  return new Paragraph({
    heading: HeadingLevel.HEADING_2,
    bidirectional: true,
    alignment: AlignmentType.RIGHT,
    spacing: { before: 220, after: 120 },
    children: [new TextRun({ text, rightToLeft: true, font: FONT, size: 26, bold: true, color: BLUE })],
  });
}
function h3(text) {
  return new Paragraph({
    heading: HeadingLevel.HEADING_3,
    bidirectional: true,
    alignment: AlignmentType.RIGHT,
    spacing: { before: 160, after: 100 },
    children: [new TextRun({ text, rightToLeft: true, font: FONT, size: 23, bold: true, color: ACCENT })],
  });
}

function bullet(text, level = 0) {
  return new Paragraph({
    bidirectional: true,
    alignment: AlignmentType.RIGHT,
    numbering: { reference: 'bullets', level },
    spacing: { after: 60, line: 288 },
    children: Array.isArray(text) ? text : [new TextRun({ text, rightToLeft: true, font: FONT, size: 22 })],
  });
}

function numItem(text, level = 0) {
  return new Paragraph({
    bidirectional: true,
    alignment: AlignmentType.RIGHT,
    numbering: { reference: 'nums', level },
    spacing: { after: 60, line: 288 },
    children: Array.isArray(text) ? text : [new TextRun({ text, rightToLeft: true, font: FONT, size: 22 })],
  });
}

function tRun(t, o = {}) {
  return new TextRun({ text: t, rightToLeft: true, font: FONT, size: o.size || 22, bold: o.bold, italics: o.italics, color: o.color });
}
// LTR code/english token
function codeRun(t, o = {}) {
  return new TextRun({ text: t, font: 'Consolas', size: o.size || 20, bold: o.bold, color: o.color || ACCENT });
}

// Generic table. cols: array of widths (dxa). header: array of strings. rows: array of arrays (string | run array)
function makeTable(cols, header, rows, opts = {}) {
  const total = cols.reduce((a, b) => a + b, 0);
  const border = { style: BorderStyle.SINGLE, size: 4, color: 'B4C2D6' };
  const borders = { top: border, bottom: border, left: border, right: border, insideHorizontal: border, insideVertical: border };

  function cell(content, width, isHeader, shade) {
    const children = Array.isArray(content)
      ? content
      : [new Paragraph({
          bidirectional: true,
          alignment: opts.align || AlignmentType.RIGHT,
          spacing: { after: 20, before: 20, line: 264 },
          children: [new TextRun({ text: String(content), rightToLeft: true, font: FONT, size: isHeader ? 20 : 19, bold: isHeader, color: isHeader ? 'FFFFFF' : '222222' })],
        })];
    return new TableCell({
      width: { size: width, type: WidthType.DXA },
      shading: { type: ShadingType.CLEAR, fill: isHeader ? HEADSHADE : (shade || 'FFFFFF'), color: 'auto' },
      margins: { top: 40, bottom: 40, left: 70, right: 70 },
      children,
    });
  }

  const headerRow = new TableRow({
    tableHeader: true,
    children: header.map((htxt, i) => cell(htxt, cols[i], true)),
  });
  const bodyRows = rows.map((r, ri) =>
    new TableRow({
      children: r.map((c, i) => cell(c, cols[i], false, ri % 2 === 1 ? GREY : 'FFFFFF')),
    })
  );

  return new Table({
    width: { size: total, type: WidthType.DXA },
    columnWidths: cols,
    visuallyRightToLeft: true,
    borders,
    rows: [headerRow, ...bodyRows],
  });
}

function spacer(after = 120) {
  return new Paragraph({ spacing: { after }, children: [new TextRun('')] });
}

// Schema field table builder: fields = [[name, type, req, desc, source]]
function schemaTable(fields) {
  const cols = [1700, 1200, 700, 3800, 1560]; // ~8960 dxa
  const header = ['שדה (Field)', 'טיפוס', 'חובה', 'תיאור', 'מקור'];
  const rows = fields.map(f => {
    const nameCell = [new Paragraph({
      bidirectional: true, alignment: AlignmentType.RIGHT, spacing: { after: 20, before: 20 },
      children: [codeRun(f[0], { size: 19, bold: true, color: NAVY })]
    })];
    const typeCell = [new Paragraph({
      bidirectional: true, alignment: AlignmentType.RIGHT, spacing: { after: 20, before: 20 },
      children: [codeRun(f[1], { size: 18, color: ACCENT })]
    })];
    return [nameCell, typeCell, f[2], f[3], f[4]];
  });
  return makeTable(cols, header, rows);
}

// ---------- Document content ----------
const children = [];

// ===== Title page =====
children.push(new Paragraph({ spacing: { before: 1800 }, children: [new TextRun('')] }));
children.push(new Paragraph({
  bidirectional: true, alignment: AlignmentType.CENTER, spacing: { after: 120 },
  children: [new TextRun({ text: 'מסמך אפיון מערכת (SRS)', rightToLeft: true, font: FONT, size: 30, bold: true, color: ACCENT })],
}));
children.push(new Paragraph({
  bidirectional: true, alignment: AlignmentType.CENTER, spacing: { after: 60 },
  children: [new TextRun({ text: 'מפה אינטראקטיבית של האימפריה הרומית', rightToLeft: true, font: FONT, size: 48, bold: true, color: NAVY })],
}));
children.push(new Paragraph({
  bidirectional: true, alignment: AlignmentType.CENTER, spacing: { after: 400 },
  children: [new TextRun({ text: 'כלכלה · דרכים · חקלאות · סחר', rightToLeft: true, font: FONT, size: 28, color: BLUE })],
}));
children.push(new Paragraph({
  bidirectional: true, alignment: AlignmentType.CENTER, spacing: { after: 100 },
  children: [new TextRun({ text: 'מסמך אפיון פונקציונלי מלא — בדגש על מבנה הנתונים והסכמות', rightToLeft: true, font: FONT, size: 22, italics: true, color: '555555' })],
}));
children.push(new Paragraph({
  bidirectional: true, alignment: AlignmentType.CENTER, spacing: { before: 1400, after: 40 },
  children: [new TextRun({ text: 'גרסה 1.0', rightToLeft: true, font: FONT, size: 22, bold: true })],
}));
children.push(new Paragraph({
  bidirectional: true, alignment: AlignmentType.CENTER,
  children: [new TextRun({ text: 'יולי 2026', rightToLeft: true, font: FONT, size: 20, color: '555555' })],
}));
children.push(new Paragraph({ children: [new PageBreak()] }));

// ===== TOC =====
children.push(h1('תוכן עניינים'));
children.push(new TableOfContents('תוכן עניינים', { hyperlink: true, headingStyleRange: '1-3' }));
children.push(new Paragraph({ children: [new PageBreak()] }));

// ===== 0. Meta table =====
children.push(h1('0. פרטי המסמך'));
children.push(makeTable(
  [2400, 6560],
  ['שדה', 'ערך'],
  [
    ['שם המערכת', 'Roman Empire Interactive Atlas (RE-Atlas)'],
    ['סוג המסמך', 'מסמך אפיון תוכנה (Software Requirements Specification)'],
    ['גרסה', '1.0'],
    ['סטטוס', 'טיוטה לאישור'],
    ['קהל יעד', 'מנהל מוצר, אדריכל מערכת, מפתחי Frontend/Backend, צוות נתונים (Data/ETL), עורכי תוכן היסטורי'],
    ['היקף גרסה נוכחית', 'אפיון פונקציונלי מלא + מודל נתונים וסכמות מפורטות'],
  ]
));
children.push(spacer());
children.push(h3('היסטוריית שינויים'));
children.push(makeTable(
  [1400, 1700, 2400, 3460],
  ['גרסה', 'תאריך', 'מחבר', 'תיאור השינוי'],
  [
    ['1.0', '22/07/2026', 'צוות אפיון', 'גרסה ראשונה — אפיון מלא על בסיס מסמך הרעיון'],
    ['', '', '', ''],
  ]
));
children.push(new Paragraph({ children: [new PageBreak()] }));

// ===== 1. Introduction =====
children.push(h1('1. מבוא ותקציר מנהלים'));

children.push(h2('1.1 מטרת המסמך'));
children.push(p('מסמך זה מגדיר את הדרישות הפונקציונליות, חוויית המשתמש, מבנה הנתונים והמגבלות של מערכת "מפה אינטראקטיבית של האימפריה הרומית". המסמך מהווה בסיס מחייב לתכנון, פיתוח, בדיקות וקבלה של המערכת, וכתוב ברמת פירוט המאפשרת לצוות פיתוח לבנות את המערכת ישירות על פיו. דגש מיוחד ניתן על מודל הנתונים והסכמות, שכן איכות המערכת נשענת על מבנה נתונים מדויק המבוסס על מקורות אקדמיים.'));

children.push(h2('1.2 רקע ומטרת המערכת'));
children.push(p('המערכת היא יישום Web המציג מפה אינטראקטיבית של האימפריה הרומית סביב שיא היקפה (בקירוב 200 לספירה). היא מאפשרת למשתמשים לחקור באופן חזותי את רשתות התחבורה (יבשה, ים ונהר), אזורי הגידול החקלאי, מוקדי המסחר וזרימת הסחורות בין הפרובינציות. המיקוד הוא בתצוגה ויזואלית ברורה, נגישות למידע היסטורי מבוסס-מחקר, ויכולת חקירה (Exploration) פשוטה באמצעות לחיצה, סינון וחיפוש.'));

children.push(h2('1.3 יעדי-על ומדדי הצלחה'));
children.push(bullet('חוויה חזותית: הצגת מפה טופוגרפית נקייה של העולם העתיק, ללא גבולות פוליטיים מודרניים, עם שכבות מידע הניתנות להדלקה/כיבוי.'));
children.push(bullet('חקירה אינטואיטיבית: כל רכיב במפה (יישוב, נתיב, אזור) לחיץ ומספק מידע מיידי ומדורג (תקציר קופץ → פאנל מלא).'));
children.push(bullet('דיוק היסטורי: כל פריט מידע ניתן לייחוס למקור אקדמי (ORBIS / DARE / OXREP).'));
children.push(bullet('ביצועים: טעינה ראשונית ראשונית ומעבר בין שכבות חלקים גם במאות/אלפי אובייקטים על המפה.'));
children.push(bullet([tRun('מדדי הצלחה מדידים: זמן טעינה ראשוני '), codeRun('< 3s'), tRun(' ; תגובת לחיצה על אובייקט '), codeRun('< 200ms'), tRun(' ; כיסוי של '), codeRun('≥ 600'), tRun(' יישובים ו-'), codeRun('≥ 1,000'), tRun(' נתיבים בגרסה הראשונה.')]));

children.push(h2('1.4 מונחון (Glossary)'));
children.push(makeTable(
  [1900, 7060],
  ['מונח', 'הגדרה'],
  [
    ['Place / Site', 'נקודת עניין גיאוגרפית: עיר, עיירה, נמל, מבצר או ציון-דרך.'],
    ['Route / Edge', 'קשת ברשת התחבורה המחברת שני צמתים; יבשתי, ימי או נהרי.'],
    ['Node', 'צומת ברשת התחבורה — לרוב מזוהה עם Place.'],
    ['Commodity', 'סחורה/תוצרת (למשל חיטה, שמן זית, זהב) הנסחרת או מיוצרת.'],
    ['Economic Zone', 'פוליגון המסמן אזור בעל מאפיין כלכלי/חקלאי דומיננטי.'],
    ['Province', 'פרובינקיה רומית מנהלית שאליה משויכים יישובים.'],
    ['Layer', 'שכבת מידע נפרדת הניתנת להדלקה/כיבוי על המפה.'],
    ['GeoJSON', 'תקן פורמט לייצוג ישויות גיאוגרפיות (נקודות, קווים, פוליגונים).'],
  ]
));
children.push(new Paragraph({ children: [new PageBreak()] }));

// ===== 2. Personas & Scope =====
children.push(h1('2. קהלי יעד, פרסונות והיקף'));

children.push(h2('2.1 פרסונות משתמש'));
children.push(makeTable(
  [1900, 2600, 4460],
  ['פרסונה', 'מאפיינים', 'צורך עיקרי מהמערכת'],
  [
    ['תלמיד / סטודנט', 'לומד היסטוריה קלאסית, ידע בסיסי', 'הבנה חזותית של רשת הסחר והגאוגרפיה; מידע קצר ונגיש.'],
    ['מורה / מרצה', 'בונה שיעור, זקוק לדוגמאות', 'הדגמת נתיבים וסחורות מול קהל; סינון לפי נושא.'],
    ['חוקר / חובב', 'ידע מעמיק, סקרנות רבה', 'חקירה חופשית, גישה למקורות ולנתונים מפורטים.'],
    ['מפתח / אינטגרטור', 'צרכן API', 'גישה למבנה נתונים אחיד וממשק תכנותי יציב.'],
  ]
));

children.push(h2('2.2 בתוך ההיקף (In Scope)'));
children.push(bullet('מפת בסיס טופוגרפית + ארבע שכבות מידע (יישובים, דרכים יבשתיות, נתיבי שיט, אזורי כלכלה/חקלאות).'));
children.push(bullet('אינטראקטיביות: לחיצה על יישוב/נתיב/אזור, תצוגת Tooltip קופצת, ופאנל פרטים מלא.'));
children.push(bullet('חיפוש טקסט חופשי וסינון לפי סחורה.'));
children.push(bullet('שכבת נתונים אחודה הנקלטת ממקורות ORBIS / DARE / OXREP באמצעות תהליך ETL.'));

children.push(h2('2.3 מחוץ להיקף (Out of Scope) — לגרסה 1'));
children.push(bullet('עריכת נתונים על-ידי משתמשי קצה (המערכת היא לקריאה בלבד — read-only).'));
children.push(bullet('חשבונות משתמש, הרשמה ושמירת מצב אישי מתמשך בשרת (מעבר לשמירה מקומית ב-localStorage).'));
children.push(bullet('סימולציית מסלול דינמית "מנקודה לנקודה" בסגנון מחשבון ORBIS (מועמד לגרסה עתידית — ראה פרק 12).'));
children.push(bullet('ציר זמן היסטורי מלא (המערכת ממקדת בחתך זמן יחיד, ~200 לספירה, עם שדות תיארוך התומכים בהרחבה עתידית).'));
children.push(new Paragraph({ children: [new PageBreak()] }));

// ===== 3. Data Sources =====
children.push(h1('3. מקורות המידע (Ground Truth)'));
children.push(p('המערכת נשענת על שלושה מאגרים אקדמיים מובילים. כל אחד ממלא תפקיד ברור במודל הנתונים, ותהליך ה-ETL (פרק 5.6) ממפה אותם לסכמה האחודה של המערכת.'));

children.push(h2('3.1 ORBIS — רשת התחבורה'));
children.push(p([tRun('The Stanford Geospatial Network Model of the Roman World. מודל גאו-מרחבי המשחזר את עלות הזמן וההוצאה הכספית של מסע בעת העתיקה, סביב 200 לספירה. הרשת מורכבת מ-'), codeRun('678'), tRun(' צמתים (nodes/places) ו-'), codeRun('1,104'), tRun(' קשתות (links) של דרכים יבשתיות, נתיבי נהר ומאות נתיבי ים בים התיכון, הים השחור והחוף האטלנטי; '), codeRun('268'), tRun(' אתרים משמשים כנמלי ים. הקשתות משוקללות לפי עלות שינוע (זמן/כסף), ולא לפי מרחק בלבד. תפקיד במערכת: מקור לשכבת הדרכים והנתיבים ולחישוב זמני המסע.')]));

children.push(h2('3.2 DARE — הבסיס הגאוגרפי'));
children.push(p([tRun('Digital Atlas of the Roman Empire (Lund University). מערכת GIS היסטורית המבוססת על Barrington Atlas ועל פרויקטי Pleiades ו-DARMC. האתרים מאורגנים כ-places ותת-אתרים (buildings/subsites), כל אחד עם מבנה נתונים ומזהה ייחודי. ההיטל הוא Spherical Mercator ('), codeRun('EPSG:3857'), tRun('), תואם לרוב מנועי מפות ה-Web, ורישוי Creative Commons BY-SA. תפקיד במערכת: מקור סמכותי למיקום, שם וסיווג של יישובים, נמלים, מבצרים ונקודות ציון, ולקישור (gazetteer) למזהי Pleiades.')]));

children.push(h2('3.3 OXREP — הכלכלה'));
children.push(p([tRun('Oxford Roman Economy Project. אוסף מסדי נתונים כלכליים הכולל, בין היתר: מסד המכרות (Mines Database, גרסה 3.0, כ-'), codeRun('1,399'), tRun(' מכרות), מסד ספינות טרופות (Shipwrecks) ומסדי בתי-בד לשמן זית וליין (presses). לכל מכרה נרשמים בין השאר: שם, אזור כרייה, מדינה, פרובינקיה, קו-אורך/רוחב, טווח תיארוך, מתכות (זהב/כסף/נחושת/עופרת/ברזל) וטכנולוגיית ניצול. תפקיד במערכת: מקור לשכבת הכלכלה — מכרות, אזורי גידול חקלאי ומוקדי ייצור.')]));

children.push(h3('3.4 סיכום מיפוי מקור → שכבה'));
children.push(makeTable(
  [2200, 3400, 3360],
  ['מקור', 'תורם למערכת', 'שכבות/ישויות'],
  [
    ['ORBIS', 'נתיבי תחבורה, זמני מסע ועלויות', 'Route, Node — שכבות דרכים ונתיבי שיט'],
    ['DARE', 'מיקום וזיהוי יישובים ונקודות ציון', 'Place — שכבת יישובים'],
    ['OXREP', 'מכרות, חקלאות ומוקדי ייצור', 'Commodity, EconomicZone — שכבת כלכלה'],
  ]
));
children.push(p([tRun('הערה: המקורות אינם משתפים מזהה אחיד. הצמדה (matching) בין ישויות תיעשה לפי מזהי '), codeRun('Pleiades'), tRun(' כשקיימים, ובאמצעות התאמה גאוגרפית וטקסטואלית כגיבוי (ראה פרק 5.6).')]));
children.push(new Paragraph({ children: [new PageBreak()] }));

// ===== 4. Functional Requirements =====
children.push(h1('4. דרישות פונקציונליות'));
children.push(p('להלן רשימת דרישות פונקציונליות ממוספרות (FR). כל דרישה בעלת מזהה ייחודי ורמת עדיפות: M = חובה (Must), S = רצוי (Should), C = יכול (Could).'));

const frCols = [1100, 6260, 1600];
const frHeader = ['מזהה', 'דרישה', 'עדיפות'];
children.push(makeTable(frCols, frHeader, [
  ['FR-01', 'המערכת תציג מפת בסיס טופוגרפית של העולם העתיק ללא גבולות פוליטיים מודרניים.', 'M'],
  ['FR-02', 'המשתמש יוכל להדליק/לכבות כל אחת מארבע שכבות המידע באופן עצמאי.', 'M'],
  ['FR-03', 'המערכת תציג יישובים כאייקונים שגודלם נגזר מדירוג החשיבות של היישוב.', 'M'],
  ['FR-04', 'המערכת תבחין ויזואלית בין סוגי נתיבים (יבשתי / ימי / נהרי) בצבע ובסגנון קו.', 'M'],
  ['FR-05', 'לחיצה על נתיב תדגיש אותו ותפתח Tooltip עם מוצא, יעד, סוג, זמן מסע וסחורות.', 'M'],
  ['FR-06', 'לחיצה על יישוב תפתח Tooltip עם שם, פרובינקיה ותוצרת מרכזית.', 'M'],
  ['FR-07', 'בחירת ישות תעדכן את פאנל הפרטים (Details Panel) עם המידע המלא.', 'M'],
  ['FR-08', 'מנוע חיפוש טקסט חופשי יאתר יישובים ופרובינקיות ויתמקד אליהם במפה.', 'M'],
  ['FR-09', 'חיפוש/בחירת סחורה יסנן את המפה ויציג רק אזורים ונתיבים רלוונטיים.', 'M'],
  ['FR-10', 'שכבת הכלכלה תוצג כפוליגונים צבועים ו/או מפת-חום לפי סוג משאב/גידול.', 'M'],
  ['FR-11', 'תפריט צד (Sidebar) יהיה ניתן להרחבה וצמצום.', 'M'],
  ['FR-12', 'פאנל הפרטים יציג מקורות (citations) לכל פריט מידע כשקיימים.', 'S'],
  ['FR-13', 'המערכת תשמור את מצב השכבות והסינון האחרון של המשתמש (localStorage).', 'S'],
  ['FR-14', 'המערכת תספק מקרא (Legend) דינמי המשתנה לפי השכבות הפעילות.', 'S'],
  ['FR-15', 'המערכת תאפשר שיתוף מצב תצוגה נוכחי באמצעות כתובת URL עם פרמטרים.', 'C'],
  ['FR-16', 'המערכת תספק ממשק API ציבורי לקריאת ישויות ונתיבים (JSON/GeoJSON).', 'S'],
]));
children.push(new Paragraph({ children: [new PageBreak()] }));

// ===== 5. DATA MODEL (emphasis) =====
children.push(h1('5. מודל הנתונים והסכמות'));
children.push(p('זהו ליבת המסמך. פרק זה מגדיר את ישויות הנתונים, השדות, הטיפוסים, היחסים ותהליך הקליטה מהמקורות. כל הישויות מיוצגות כ-JSON, וישויות גאוגרפיות נשמרות ומוגשות גם בפורמט GeoJSON.'));

children.push(h2('5.1 סקירת מודל הנתונים (ERD מילולי)'));
children.push(bullet([codeRun('Place'), tRun(' — יישוב/נקודת עניין. משויך ל-'), codeRun('Province'), tRun(' אחת, ומקושר ל-0..N '), codeRun('Commodity'), tRun(' (תוצרת).')]));
children.push(bullet([codeRun('Route'), tRun(' — נתיב בין שני '), codeRun('Place'), tRun(' (origin/destination). נושא רשימת '), codeRun('Commodity'), tRun(' וזמני מסע.')]));
children.push(bullet([codeRun('Commodity'), tRun(' — סחורה. מקושרת ל-'), codeRun('Place'), tRun(' (ייצור/סחר) ול-'), codeRun('Route'), tRun(' (מטען) ול-'), codeRun('EconomicZone'), tRun('.')]));
children.push(bullet([codeRun('EconomicZone'), tRun(' — פוליגון אזור כלכלי, מכיל '), codeRun('Commodity'), tRun(' דומיננטית אחת או יותר.')]));
children.push(bullet([codeRun('Province'), tRun(' — פרובינקיה, מכילה 0..N '), codeRun('Place'), tRun('.')]));
children.push(bullet([codeRun('Source'), tRun(' — רשומת מקור/ציטוט המקושרת מכל ישות אחרת לצורך ייחוס (provenance).')]));

children.push(h2('5.2 ישות: Place (יישוב / נקודת עניין)'));
children.push(p('מייצג עיר, עיירה, נמל, מבצר או ציון-דרך. הבסיס הגאוגרפי מגיע מ-DARE; שדות כלכליים מ-OXREP; קישוריות מ-ORBIS.'));
children.push(schemaTable([
  ['id', 'string', 'כן', 'מזהה פנימי ייחודי (UUID/slug), למשל "place_rome".', 'פנימי'],
  ['name', 'string', 'כן', 'שם באנגלית (Rome, Alexandria).', 'DARE'],
  ['nameLatin', 'string', 'לא', 'שם לטיני היסטורי (Roma).', 'DARE'],
  ['nameHe', 'string', 'לא', 'שם בעברית לתצוגה מתורגמת.', 'עריכה'],
  ['type', 'enum', 'כן', 'city | town | port | fort | landmark.', 'DARE'],
  ['importance', 'integer', 'כן', 'דירוג 1–5 לקביעת גודל אייקון.', 'נגזר'],
  ['coordinates', 'number[2]', 'כן', '[lon, lat] ב-WGS84.', 'DARE'],
  ['provinceId', 'string', 'כן', 'מפתח זר ל-Province.', 'DARE'],
  ['pleiadesId', 'string', 'לא', 'מזהה Pleiades לקישור חיצוני והצמדה.', 'DARE'],
  ['isPort', 'boolean', 'כן', 'האם משמש כנמל ים (ב-ORBIS).', 'ORBIS'],
  ['commodityIds', 'string[]', 'לא', 'רשימת סחורות המיוצרות/נסחרות במקום.', 'OXREP'],
  ['description', 'string', 'לא', 'תיאור היסטורי קצר לפאנל הפרטים.', 'עריכה'],
  ['sourceIds', 'string[]', 'לא', 'מפתחות זרים ל-Source (ייחוס).', 'פנימי'],
]));

children.push(h2('5.3 ישות: Route (נתיב תחבורה)'));
children.push(p('קשת ברשת התחבורה. המקור העיקרי הוא ORBIS, כולל זמני מסע ועלויות. הגיאומטריה נשמרת כ-LineString.'));
children.push(schemaTable([
  ['id', 'string', 'כן', 'מזהה ייחודי, למשל "route_ostia_alexandria".', 'פנימי'],
  ['originId', 'string', 'כן', 'מפתח זר ל-Place (מוצא).', 'ORBIS'],
  ['destinationId', 'string', 'כן', 'מפתח זר ל-Place (יעד).', 'ORBIS'],
  ['mode', 'enum', 'כן', 'road | sea | river.', 'ORBIS'],
  ['geometry', 'LineString', 'כן', 'מערך נקודות [lon,lat] של מסלול הקו.', 'ORBIS'],
  ['distanceKm', 'number', 'לא', 'אורך משוער בק"מ.', 'ORBIS'],
  ['travelDays', 'number', 'כן', 'זמן מסע משוער בימים (בתנאים טיפוסיים).', 'ORBIS'],
  ['travelSeason', 'enum', 'לא', 'summer | winter — משפיע על זמן ים.', 'ORBIS'],
  ['costDenarii', 'number', 'לא', 'עלות שינוע יחסית (מודל ORBIS).', 'ORBIS'],
  ['commodityIds', 'string[]', 'לא', 'סחורות עיקריות שעברו בנתיב.', 'OXREP/עריכה'],
  ['sourceIds', 'string[]', 'לא', 'ייחוס למקורות.', 'פנימי'],
]));

children.push(h2('5.4 ישות: Commodity (סחורה)'));
children.push(schemaTable([
  ['id', 'string', 'כן', 'מזהה, למשל "olive_oil".', 'פנימי'],
  ['name', 'string', 'כן', 'שם באנגלית (Olive Oil, Gold, Grain).', 'OXREP'],
  ['nameHe', 'string', 'לא', 'שם בעברית (שמן זית, זהב, תבואה).', 'עריכה'],
  ['category', 'enum', 'כן', 'agriculture | metal | mineral | manufactured | livestock.', 'OXREP'],
  ['icon', 'string', 'לא', 'מזהה אייקון לתצוגה במקרא/סינון.', 'עיצוב'],
  ['color', 'string', 'לא', 'צבע Hex לצביעת שכבת הכלכלה.', 'עיצוב'],
  ['description', 'string', 'לא', 'תיאור קצר של הסחורה והקשרה הכלכלי.', 'עריכה'],
]));

children.push(h2('5.5 ישות: EconomicZone (אזור כלכלי/חקלאי)'));
children.push(p('פוליגון גאוגרפי המסמן אזור בעל מאפיין כלכלי דומיננטי (למשל מצרים כ"אסם התבואה", ספרד כיצרנית שמן ומתכות). מוצג כשכבת פוליגונים צבועים או מפת-חום.'));
children.push(schemaTable([
  ['id', 'string', 'כן', 'מזהה, למשל "zone_egypt_grain".', 'פנימי'],
  ['name', 'string', 'כן', 'שם האזור לתצוגה.', 'OXREP/עריכה'],
  ['geometry', 'Polygon', 'כן', 'פוליגון גבולות האזור (GeoJSON).', 'עריכה'],
  ['dominantCommodityIds', 'string[]', 'כן', 'סחורות מאפיינות את האזור.', 'OXREP'],
  ['intensity', 'number', 'לא', 'עצמה 0–1 למפת-חום/שקיפות.', 'נגזר'],
  ['description', 'string', 'לא', 'הקשר כלכלי-היסטורי של האזור.', 'עריכה'],
  ['sourceIds', 'string[]', 'לא', 'ייחוס למקורות.', 'פנימי'],
]));

children.push(h2('5.6 ישויות תומכות: Province ו-Source'));
children.push(h3('Province'));
children.push(schemaTable([
  ['id', 'string', 'כן', 'מזהה, למשל "prov_aegyptus".', 'פנימי'],
  ['name', 'string', 'כן', 'שם הפרובינקיה (Aegyptus, Hispania).', 'DARE'],
  ['nameHe', 'string', 'לא', 'שם בעברית.', 'עריכה'],
  ['geometry', 'Polygon', 'לא', 'גבול משוער (לתצוגה/סינון בלבד).', 'עריכה'],
]));
children.push(h3('Source (ייחוס / Provenance)'));
children.push(schemaTable([
  ['id', 'string', 'כן', 'מזהה מקור.', 'פנימי'],
  ['label', 'string', 'כן', 'תיאור קצר (מחבר/מאגר).', 'פנימי'],
  ['dataset', 'enum', 'כן', 'ORBIS | DARE | OXREP | other.', 'פנימי'],
  ['url', 'string', 'לא', 'קישור למקור המקוון.', 'פנימי'],
  ['license', 'string', 'לא', 'רישיון (למשל CC BY-SA).', 'פנימי'],
]));

children.push(h2('5.7 תהליך קליטת נתונים (ETL) ו-Provenance'));
children.push(p('הנתונים נקלטים מהמקורות פעם אחת (או בעדכונים תקופתיים) ומנורמלים לסכמה האחודה. התהליך מתועד כדי לשמר שקיפות מדעית ולאפשר עדכונים.'));
children.push(numItem([tRun('Extract — ייבוא גולמי: ORBIS (רשת nodes/links), DARE (places, GeoJSON, מזהי Pleiades), OXREP (טבלאות Excel/CSV של מכרות, ספינות ובתי-בד).')]));
children.push(numItem([tRun('Match — הצמדת ישויות בין המקורות לפי '), codeRun('pleiadesId'), tRun(' כשקיים; אחרת התאמה גאוגרפית (רדיוס סף) + התאמת שם מנורמל.')]));
children.push(numItem('Transform — נרמול לשמות שדות אחידים, המרת יחידות, חישוב importance ו-intensity, תרגום שדות תצוגה לעברית.'));
children.push(numItem([tRun('Load — כתיבה למאגר בפורמט GeoJSON/JSON, בתוספת שדה '), codeRun('sourceIds'), tRun(' לכל רשומה לצורך ייחוס.')]));
children.push(numItem('Validate — בדיקת שלמות: מפתחות זרים קיימים, קואורדינטות בטווח, אין נתיב ללא מוצא/יעד (ראה פרק 9).'));
children.push(p([tRun('כלל מנחה: כל רשומה מוצגת חייבת להיות ניתנת לייחוס למקור אחד לפחות. שדה '), codeRun('sourceIds'), tRun(' אינו ריק עבור ישויות תוכן היסטורי.')]));
children.push(new Paragraph({ children: [new PageBreak()] }));

// ===== 5.8 JSON examples =====
children.push(h2('5.8 דוגמאות JSON'));
children.push(p('דוגמת רשומת Place (GeoJSON Feature):'));
const codeShade = 'F5F2EA';
function codeBlock(lines) {
  return lines.map((ln, i) => new Paragraph({
    shading: { type: ShadingType.CLEAR, fill: codeShade, color: 'auto' },
    spacing: { after: i === lines.length - 1 ? 140 : 0, before: i === 0 ? 60 : 0, line: 240 },
    indent: { left: 120, right: 120 },
    children: [new TextRun({ text: ln, font: 'Consolas', size: 18, color: '333333' })],
  }));
}
codeBlock([
  '{',
  '  "type": "Feature",',
  '  "id": "place_alexandria",',
  '  "geometry": { "type": "Point", "coordinates": [29.9187, 31.2001] },',
  '  "properties": {',
  '    "name": "Alexandria", "nameLatin": "Alexandria", "nameHe": "אלכסנדריה",',
  '    "type": "port", "importance": 5, "isPort": true,',
  '    "provinceId": "prov_aegyptus", "pleiadesId": "727070",',
  '    "commodityIds": ["grain", "papyrus", "glass"],',
  '    "sourceIds": ["src_dare", "src_oxrep"]',
  '  }',
  '}',
].map(x=>x)).forEach(pp => children.push(pp));

children.push(p('דוגמת רשומת Route:'));
codeBlock([
  '{',
  '  "id": "route_alexandria_ostia",',
  '  "originId": "place_alexandria", "destinationId": "place_ostia",',
  '  "mode": "sea", "travelDays": 14, "travelSeason": "summer",',
  '  "distanceKm": 1900, "costDenarii": 24.5,',
  '  "commodityIds": ["grain", "papyrus"],',
  '  "geometry": { "type": "LineString",',
  '    "coordinates": [[29.9,31.2],[24.1,35.5],[12.3,41.7]] },',
  '  "sourceIds": ["src_orbis"]',
  '}',
]).forEach(pp => children.push(pp));
children.push(new Paragraph({ children: [new PageBreak()] }));

// ===== 6. Map & Layers =====
children.push(h1('6. מבנה המפה ושכבות המידע'));
children.push(h2('6.1 מפת בסיס'));
children.push(p([tRun('מפת בסיס טופוגרפית של העולם העתיק, ללא גבולות פוליטיים מודרניים. היטל '), codeRun('EPSG:3857'), tRun(' (Spherical Mercator) לתאימות עם מנועי מפות Web. אזור ברירת המחדל: אגן הים התיכון. רמות זום נתמכות: כלל-אימפריה עד רמת עיר בודדת.')]));

children.push(h2('6.2 שכבות המידע'));
children.push(makeTable(
  [1700, 2500, 4760],
  ['שכבה', 'ייצוג ויזואלי', 'תוכן ומקור'],
  [
    ['יישובים', 'אייקונים בגודל משתנה', 'ערים, עיירות ונמלים; גודל לפי importance. מקור: DARE.'],
    ['דרכים יבשתיות', 'קווים מלאים (Via)', 'רשת הדרכים הסלולות. מקור: ORBIS (mode=road).'],
    ['נתיבי שיט', 'קווים מקווקווים', 'סחר ימי בים התיכון/השחור ונהרות (ריין, נילוס). ORBIS (sea/river).'],
    ['כלכלה וחקלאות', 'פוליגונים צבועים / מפת-חום', 'גידולים ומשאבים לפי אזור. מקור: OXREP.'],
  ]
));

children.push(h2('6.3 מקרא (Legend) ובקרת שכבות'));
children.push(bullet('כל שכבה ניתנת להדלקה/כיבוי בנפרד מתוך בקרת שכבות בתפריט הצד.'));
children.push(bullet('המקרא דינמי: מציג רק את הסמלים והצבעים של השכבות הפעילות כרגע.'));
children.push(bullet('בעת סינון לפי סחורה, המקרא מדגיש את הצבע/סמל של אותה סחורה.'));

children.push(h2('6.4 התנהגות לפי רמת זום (Clustering)'));
children.push(bullet([tRun('ברמת זום נמוכה: הצגת יישובים בעלי '), codeRun('importance ≥ 4'), tRun(' בלבד, ואשכולות (clusters) לשאר.')]));
children.push(bullet('בהגדלה: חשיפה הדרגתית של יישובים קטנים יותר וספירת אשכולות מתעדכנת.'));
children.push(bullet('נתיבים דחוסים עשויים להתמזג ויזואלית ברמת זום נמוכה, ולהיפרד בהגדלה.'));
children.push(new Paragraph({ children: [new PageBreak()] }));

// ===== 7. Interactivity =====
children.push(h1('7. אינטראקטיביות וחוויית משתמש'));
children.push(p('עקרון מנחה: מידע מדורג. שלב ראשון — Tooltip קופץ תמציתי; שלב שני — פאנל פרטים מלא. כל רכיב במפה לחיץ.'));

children.push(h2('7.1 לחיצה על נתיב (Route Click)'));
children.push(bullet('הנתיב הנבחר מודגש ויזואלית (עובי/צבע) ושאר המפה מעומעמת מעט (dim).'));
children.push(bullet('נפתחת חלונית Tooltip צפה ליד הנתיב עם: נקודת מוצא ויעד; סוג הנתיב (יבשתי/ימי/נהרי); זמן מסע משוער (למשל "14 ימים בהפלגה"); סחורות עיקריות (למשל תבואה, פפירוס, יין).'));
children.push(bullet('פאנל הפרטים מתעדכן במקביל עם המידע המלא של הנתיב.'));

children.push(h2('7.2 לחיצה על יישוב/אזור (Location Click)'));
children.push(bullet('Tooltip מציג: שם המקום, שיוך לפרובינקיה, ותוצרת מרכזית.'));
children.push(bullet('בחירת אזור כלכלי מציגה את הסחורות הדומיננטיות שלו ואת עצמת הפעילות.'));
children.push(bullet('פאנל הפרטים מציג תיאור, הקשר היסטורי ורשימת מקורות.'));

children.push(h2('7.3 מצבי אינטראקציה (Hover / Select / Filter)'));
children.push(makeTable(
  [1700, 3630, 3630],
  ['מצב', 'טריגר', 'תגובת המערכת'],
  [
    ['Hover', 'מעבר עכבר מעל אובייקט', 'הדגשה קלה + Tooltip מינימלי (שם בלבד).'],
    ['Select', 'לחיצה על אובייקט', 'הדגשה מלאה + Tooltip מפורט + עדכון פאנל.'],
    ['Filter', 'בחירת סחורה/חיפוש', 'עמעום לא-רלוונטיים; הבלטת תואמים בלבד.'],
    ['Deselect', 'לחיצה על רקע/Esc', 'איפוס הדגשות; סגירת Tooltip.'],
  ]
));

children.push(h2('7.4 נגישות ורספונסיביות'));
children.push(bullet('תמיכה מלאה ב-RTL בממשק העברי; טקסטים לטיניים (שמות אתרים) מוצגים LTR בתוך הקשר RTL.'));
children.push(bullet('ניווט מקלדת: Tab בין רכיבי הבקרה, Enter לבחירה, Esc לסגירה.'));
children.push(bullet('עמידה בהנחיות WCAG 2.1 AA: ניגודיות צבעים, טקסט חלופי לאייקונים, יעדי לחיצה מספקים.'));
children.push(bullet('פריסה רספונסיבית: במסך צר תפריט הצד הופך למגירה נשלפת (drawer) והמפה תופסת את מרב השטח.'));
children.push(new Paragraph({ children: [new PageBreak()] }));

// ===== 8. Sidebar & Search =====
children.push(h1('8. תפריט צד, חיפוש ופאנל פרטים'));

children.push(h2('8.1 תפריט צד (Sidebar)'));
children.push(bullet('ניתן להרחבה וצמצום; במצב מצומצם מוצגים אייקונים בלבד.'));
children.push(bullet('מכיל: מנוע חיפוש, בקרת שכבות, מסנני סחורה, ופאנל פרטים.'));

children.push(h2('8.2 מנוע חיפוש (Search)'));
children.push(bullet('חיפוש טקסט חופשי בשמות יישובים ופרובינקיות (למשל "Ephesus") — בחירה ממקדת את המפה אל התוצאה.'));
children.push(bullet('חיפוש/בחירת סחורה (למשל "Olive Oil" או "Gold") — מסנן את המפה ומציג רק אזורים ונתיבים רלוונטיים.'));
children.push(bullet('הצעות אוטומטיות (autocomplete) תוך כדי הקלדה, עם סימון סוג התוצאה (יישוב / פרובינקיה / סחורה).'));
children.push(bullet('חיפוש תומך בעברית ובאנגלית ומתעלם מרישיות וניקוד.'));

children.push(h2('8.3 פאנל פרטים (Details Panel)'));
children.push(p('כאשר ישות נבחרת במפה, הפאנל מציג את מלוא המידע — מעבר לתקציר שב-Tooltip.'));
children.push(makeTable(
  [2200, 6760],
  ['ישות נבחרת', 'תוכן הפאנל'],
  [
    ['יישוב', 'שם (עברית/לטינית), סוג, פרובינקיה, קואורדינטות, תוצרת, תיאור והקשר היסטורי, מקורות.'],
    ['נתיב', 'מוצא ויעד, סוג, מרחק, זמן מסע (לפי עונה), עלות יחסית, סחורות, מקורות.'],
    ['אזור כלכלי', 'שם, סחורות דומיננטיות, עצמה, הקשר כלכלי-היסטורי, מקורות.'],
  ]
));
children.push(new Paragraph({ children: [new PageBreak()] }));

// ===== 9. API =====
children.push(h1('9. ממשק תכנותי (API) וחוזה נתונים'));
children.push(p('ה-Frontend צורך נתונים משכבת שירות (REST). כל התשובות בפורמט JSON; ישויות גאוגרפיות כ-GeoJSON FeatureCollection. פרק זה מגדיר את החוזה בין השרת ל-Client.'));
children.push(makeTable(
  [2900, 900, 5160],
  ['Endpoint', 'Method', 'תיאור ותשובה'],
  [
    ['/api/places', 'GET', 'רשימת יישובים (GeoJSON). פרמטרים: province, type, bbox, commodity.'],
    ['/api/places/{id}', 'GET', 'פרטי יישוב מלאים כולל מקורות.'],
    ['/api/routes', 'GET', 'רשימת נתיבים. פרמטרים: mode, commodity, origin, destination.'],
    ['/api/routes/{id}', 'GET', 'פרטי נתיב מלאים כולל זמני מסע וסחורות.'],
    ['/api/commodities', 'GET', 'קטלוג הסחורות (לסינון ולמקרא).'],
    ['/api/zones', 'GET', 'אזורים כלכליים (GeoJSON Polygons). פרמטר: commodity.'],
    ['/api/search', 'GET', 'חיפוש מאוחד. פרמטר: q — מחזיר יישובים/פרובינקיות/סחורות.'],
  ]
));
children.push(spacer());
children.push(h3('עקרונות חוזה'));
children.push(bullet([tRun('כל ישות מחזירה '), codeRun('id'), tRun(' יציב; מפתחות זרים ('), codeRun('provinceId'), tRun(', '), codeRun('commodityIds'), tRun(') מפנים לישויות אחרות.')]));
children.push(bullet([tRun('סינון בצד השרת מפחית עומס: '), codeRun('/api/routes?commodity=olive_oil'), tRun(' מחזיר רק נתיבים רלוונטיים.')]));
children.push(bullet([tRun('גרסאות API בנתיב ('), codeRun('/api/v1/'), tRun(') לשמירת תאימות לאחור.')]));
children.push(new Paragraph({ children: [new PageBreak()] }));

// ===== 10. NFR =====
children.push(h1('10. דרישות לא-פונקציונליות (NFR)'));
children.push(makeTable(
  [2200, 1400, 5360],
  ['תחום', 'מזהה', 'דרישה'],
  [
    ['ביצועים', 'NFR-01', 'טעינה ראשונית של המפה עד 3 שניות בחיבור סביר.'],
    ['ביצועים', 'NFR-02', 'תגובת לחיצה/בחירה על אובייקט עד 200ms.'],
    ['ביצועים', 'NFR-03', 'רינדור חלק של אלפי אובייקטים באמצעות clustering ו-vector tiles.'],
    ['נגישות', 'NFR-04', 'עמידה ב-WCAG 2.1 AA כולל ניווט מקלדת ותמיכת קורא מסך.'],
    ['בינאום', 'NFR-05', 'תמיכת RTL מלאה (עברית) עם טיפול נכון בטקסט לטיני מעורב.'],
    ['תאימות', 'NFR-06', 'תמיכה בדפדפנים מודרניים: Chrome, Firefox, Safari, Edge (2 גרסאות אחרונות).'],
    ['רספונסיביות', 'NFR-07', 'תצוגה תקינה במסכי דסקטופ, טאבלט ומובייל.'],
    ['אמינות נתונים', 'NFR-08', 'כל פריט תוכן היסטורי ניתן לייחוס למקור (provenance).'],
    ['תחזוקתיות', 'NFR-09', 'הפרדה בין שכבת נתונים, שירות ו-UI; נתונים כקבצים/DB הניתנים לעדכון.'],
    ['אבטחה', 'NFR-10', 'API לקריאה בלבד; הגנת CORS ו-rate limiting בסיסי.'],
  ]
));
children.push(new Paragraph({ children: [new PageBreak()] }));

// ===== 11. Architecture (light) =====
children.push(h1('11. ארכיטקטורה ברמה גבוהה'));
children.push(p('הארכיטקטורה מודולרית ומפרידה בין שלוש שכבות. הדגש הוא על שכבת הנתונים, שהיא לב המערכת.'));
children.push(h2('11.1 שכבות המערכת'));
children.push(numItem([tRun('שכבת נתונים (Data Layer): מאגר GeoJSON/JSON או בסיס נתונים מרחבי (PostGIS), עם תוצרי ה-ETL מ-ORBIS/DARE/OXREP.')]));
children.push(numItem([tRun('שכבת שירות (API Layer): שירות REST המגיש ישויות מסוננות (פרק 9), עם מטמון (cache) לתשובות נפוצות.')]));
children.push(numItem([tRun('שכבת הצגה (Client): יישום Web עם מנוע מפות (למשל Leaflet / MapLibre GL), בקרת שכבות, Tooltip ופאנל פרטים.')]));
children.push(h2('11.2 המלצות טכנולוגיה (לא מחייב)'));
children.push(bullet([tRun('מפות: '), codeRun('MapLibre GL JS'), tRun(' או '), codeRun('Leaflet'), tRun(' — תמיכה ב-vector tiles ל-clustering יעיל.')]));
children.push(bullet([tRun('פורמט נתונים מרחבי: '), codeRun('GeoJSON'), tRun(' לפיתוח, '), codeRun('vector tiles (PMTiles/MBTiles)'), tRun(' לקנה מידה.')]));
children.push(bullet([tRun('מאגר: קבצים סטטיים לגרסת MVP, או '), codeRun('PostgreSQL + PostGIS'), tRun(' לשאילתות מרחביות מתקדמות.')]));
children.push(bullet([tRun('Frontend: מסגרת קומפוננטות (React/Vue) עם ניהול מצב לשכבות ולסינון; '), codeRun('localStorage'), tRun(' לשמירת העדפות.')]));

children.push(h2('11.3 זרימת נתונים טיפוסית'));
children.push(numItem('המשתמש מדליק שכבה או מבצע חיפוש/סינון.'));
children.push(numItem('ה-Client שולח בקשה ל-API עם פרמטרי סינון.'));
children.push(numItem('ה-API מחזיר GeoJSON מסונן; ה-Client מרנדר על המפה.'));
children.push(numItem('לחיצה על אובייקט טוענת פרטים מלאים (או משתמשת בנתונים שכבר נטענו) לפאנל.'));
children.push(new Paragraph({ children: [new PageBreak()] }));

// ===== 12. Roadmap =====
children.push(h1('12. אבני דרך ותכולת גרסאות'));
children.push(makeTable(
  [1500, 2300, 5160],
  ['שלב', 'כותרת', 'תכולה'],
  [
    ['Phase 0', 'תשתית נתונים', 'ETL מ-ORBIS/DARE/OXREP; בניית הסכמה האחודה; אימות נתונים.'],
    ['Phase 1 (MVP)', 'מפה ושכבות', 'מפת בסיס, שכבת יישובים ודרכים, Tooltip בסיסי, בקרת שכבות.'],
    ['Phase 2', 'אינטראקציה מלאה', 'נתיבי שיט, שכבת כלכלה, פאנל פרטים מלא, חיפוש וסינון סחורה.'],
    ['Phase 3', 'העשרה', 'מקורות/ציטוטים, שיתוף URL, שיפורי נגישות וביצועים.'],
    ['Phase 4 (עתידי)', 'הרחבות', 'מחשבון מסלול דינמי (ORBIS-like), ציר זמן היסטורי, שכבות נוספות.'],
  ]
));

children.push(h2('12.1 קריטריוני קבלה ל-MVP (Phase 1)'));
children.push(bullet('המפה נטענת ומציגה יישובים ודרכים יבשתיות עם בקרת שכבות תקינה.'));
children.push(bullet('לחיצה על יישוב/נתיב פותחת Tooltip עם המידע התמציתי הנדרש.'));
children.push(bullet('גודל אייקון היישוב משקף את דירוג החשיבות.'));
children.push(bullet('הנתונים נטענים מ-API בפורמט GeoJSON התואם לסכמה בפרק 5.'));
children.push(new Paragraph({ children: [new PageBreak()] }));

// ===== 13. Risks & assumptions =====
children.push(h1('13. סיכונים, הנחות ותלויות'));
children.push(h2('13.1 הנחות'));
children.push(bullet('הנתונים מהמקורות זמינים להורדה ושימוש בכפוף לרישיונם (למשל DARE — CC BY-SA).'));
children.push(bullet('חתך הזמן המרכזי הוא ~200 לספירה, בהתאם ל-ORBIS.'));
children.push(bullet('היקף התוכן ההיסטורי המילולי (תיאורים) יוזן על-ידי עורך תוכן, ולא נגזר אוטומטית.'));

children.push(h2('13.2 סיכונים'));
children.push(makeTable(
  [3200, 1400, 4360],
  ['סיכון', 'חומרה', 'מיטיגציה'],
  [
    ['אי-התאמה בין מזהי המקורות (matching)', 'גבוהה', 'שימוש ב-Pleiades כמפתח-על + התאמה גאוגרפית וסקירה ידנית של מקרי קצה.'],
    ['ביצועים עם אלפי אובייקטים', 'בינונית', 'Clustering, vector tiles וסינון בצד שרת.'],
    ['פערים/חוסרים בנתוני סחורות לכל נתיב', 'בינונית', 'סימון שדות חסרים; השלמה עריכתית מבוססת מקורות.'],
    ['דיוק היסטורי שנוי במחלוקת', 'נמוכה', 'הצגת מקורות שקופה לכל פריט; ניסוח זהיר ("משוער").'],
  ]
));

children.push(h2('13.3 תלויות'));
children.push(bullet('זמינות והרשאות שימוש בנתוני ORBIS, DARE ו-OXREP.'));
children.push(bullet('מנוע מפות צד-שלישי ואריחי בסיס (base tiles) מתאימים לעולם העתיק.'));
children.push(new Paragraph({ children: [new PageBreak()] }));

// ===== 14. Appendix sources =====
children.push(h1('14. נספח: מקורות אקדמיים'));
children.push(p('המקורות הבאים משמשים כ-Ground Truth למערכת ומצוטטים בשדות ה-Source:'));

function linkPara(label, url) {
  return new Paragraph({
    bidirectional: true, alignment: AlignmentType.RIGHT, spacing: { after: 80, line: 288 },
    numbering: { reference: 'bullets', level: 0 },
    children: [
      tRun(label + ' — '),
      new ExternalHyperlink({ link: url, children: [new TextRun({ text: url, font: 'Consolas', size: 18, color: '2E5496', underline: {} })] }),
    ],
  });
}
children.push(linkPara('ORBIS: The Stanford Geospatial Network Model of the Roman World', 'https://orbis.stanford.edu'));
children.push(linkPara('DARE: Digital Atlas of the Roman Empire (Lund University)', 'https://dh.gu.se/dare/'));
children.push(linkPara('OXREP: The Oxford Roman Economy Project', 'https://oxrep.web.ox.ac.uk'));
children.push(linkPara('OXREP Mines Database', 'https://oxrep.web.ox.ac.uk/mines-database'));
children.push(linkPara('Pleiades gazetteer (קישור מזהים)', 'https://pleiades.stoa.org'));
children.push(spacer(200));
children.push(p([tRun('— סוף המסמך —')], { align: AlignmentType.CENTER, color: '888888', italics: true }));

// ---------- Assemble ----------
const doc = new Document({
  creator: 'RE-Atlas',
  title: 'מסמך אפיון — מפה אינטראקטיבית של האימפריה הרומית',
  styles: {
    default: {
      document: { run: { font: FONT, size: 22 } },
    },
  },
  numbering: {
    config: [
      {
        reference: 'bullets',
        levels: [
          { level: 0, format: LevelFormat.BULLET, text: '•', alignment: AlignmentType.LEFT, style: { run: { color: ACCENT }, paragraph: { indent: { right: 360, hanging: 220 } } } },
          { level: 1, format: LevelFormat.BULLET, text: '◦', alignment: AlignmentType.LEFT, style: { paragraph: { indent: { right: 720, hanging: 220 } } } },
        ],
      },
      {
        reference: 'nums',
        levels: [
          { level: 0, format: LevelFormat.DECIMAL, text: '%1.', alignment: AlignmentType.RIGHT, style: { run: { bold: true, color: NAVY }, paragraph: { indent: { right: 360, hanging: 260 } } } },
        ],
      },
    ],
  },
  sections: [{
    properties: {
      page: {
        size: { width: 11906, height: 16838 }, // A4
        margin: { top: 1200, bottom: 1200, left: 1100, right: 1100 },
      },
    },
    headers: {
      default: new Header({
        children: [new Paragraph({
          bidirectional: true, alignment: AlignmentType.LEFT,
          border: { bottom: { style: BorderStyle.SINGLE, size: 4, color: 'CCCCCC', space: 4 } },
          children: [new TextRun({ text: 'RE-Atlas · מסמך אפיון v1.0', rightToLeft: true, font: FONT, size: 16, color: '888888' })],
        })],
      }),
    },
    footers: {
      default: new Footer({
        children: [new Paragraph({
          alignment: AlignmentType.CENTER,
          children: [
            new TextRun({ text: 'עמוד ', rightToLeft: true, font: FONT, size: 16, color: '888888' }),
            new TextRun({ children: [PageNumber.CURRENT], font: FONT, size: 16, color: '888888' }),
          ],
        })],
      }),
    },
    children,
  }],
});

Packer.toBuffer(doc).then(buf => {
  fs.writeFileSync(process.argv[2] || 'output.docx', buf);
  console.log('WROTE', (process.argv[2] || 'output.docx'), buf.length, 'bytes');
});
