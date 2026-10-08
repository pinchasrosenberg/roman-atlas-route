const {
  Document, Packer, Paragraph, TextRun, HeadingLevel, AlignmentType,
  Table, TableRow, TableCell, WidthType, BorderStyle, ShadingType,
  TableOfContents, PageBreak, Header, Footer, PageNumber, LevelFormat, ExternalHyperlink
} = require('docx');
const fs = require('fs');

const FONT='Arial', NAVY='1F3864', BLUE='2E5496', ACCENT='A6611A', GREY='F2F2F2', HEAD='1F3864';

function p(text, o={}) {
  const runs = Array.isArray(text) ? text : [new TextRun({ text, rightToLeft:true, font:FONT, size:o.size||22, bold:o.bold, italics:o.italics, color:o.color })];
  return new Paragraph({ bidirectional:true, alignment:o.align||AlignmentType.RIGHT, spacing:{ after:o.after!=null?o.after:120, line:300, before:o.before||0 }, children:runs });
}
function h1(t){return new Paragraph({heading:HeadingLevel.HEADING_1,bidirectional:true,alignment:AlignmentType.RIGHT,spacing:{before:300,after:150},children:[new TextRun({text:t,rightToLeft:true,font:FONT,size:32,bold:true,color:NAVY})]});}
function h2(t){return new Paragraph({heading:HeadingLevel.HEADING_2,bidirectional:true,alignment:AlignmentType.RIGHT,spacing:{before:220,after:110},children:[new TextRun({text:t,rightToLeft:true,font:FONT,size:26,bold:true,color:BLUE})]});}
function h3(t){return new Paragraph({heading:HeadingLevel.HEADING_3,bidirectional:true,alignment:AlignmentType.RIGHT,spacing:{before:150,after:90},children:[new TextRun({text:t,rightToLeft:true,font:FONT,size:23,bold:true,color:ACCENT})]});}
function tR(t,o={}){return new TextRun({text:t,rightToLeft:true,font:FONT,size:o.size||22,bold:o.bold,italics:o.italics,color:o.color});}
function code(t,o={}){return new TextRun({text:t,font:'Consolas',size:o.size||20,bold:o.bold,color:o.color||ACCENT});}
function bullet(text,lvl=0){return new Paragraph({bidirectional:true,alignment:AlignmentType.RIGHT,numbering:{reference:'b',level:lvl},spacing:{after:60,line:288},children:Array.isArray(text)?text:[new TextRun({text,rightToLeft:true,font:FONT,size:22})]});}
function num(text,lvl=0){return new Paragraph({bidirectional:true,alignment:AlignmentType.RIGHT,numbering:{reference:'n',level:lvl},spacing:{after:60,line:288},children:Array.isArray(text)?text:[new TextRun({text,rightToLeft:true,font:FONT,size:22})]});}
function spacer(a=120){return new Paragraph({spacing:{after:a},children:[new TextRun('')]});}

function table(cols,header,rows,opts={}){
  const total=cols.reduce((a,b)=>a+b,0);
  const bd={style:BorderStyle.SINGLE,size:4,color:'B4C2D6'};
  const borders={top:bd,bottom:bd,left:bd,right:bd,insideHorizontal:bd,insideVertical:bd};
  function cell(content,w,isH,shade){
    const children=Array.isArray(content)?content:[new Paragraph({bidirectional:true,alignment:opts.align||AlignmentType.RIGHT,spacing:{after:20,before:20,line:264},children:[new TextRun({text:String(content),rightToLeft:true,font:FONT,size:isH?20:18,bold:isH,color:isH?'FFFFFF':'222222'})]})];
    return new TableCell({width:{size:w,type:WidthType.DXA},shading:{type:ShadingType.CLEAR,fill:isH?HEAD:(shade||'FFFFFF'),color:'auto'},margins:{top:40,bottom:40,left:70,right:70},children});
  }
  const hr=new TableRow({tableHeader:true,children:header.map((t,i)=>cell(t,cols[i],true))});
  const br=rows.map((r,ri)=>new TableRow({children:r.map((c,i)=>cell(c,cols[i],false,ri%2?GREY:'FFFFFF'))}));
  return new Table({width:{size:total,type:WidthType.DXA},columnWidths:cols,visuallyRightToLeft:true,borders,rows:[hr,...br]});
}
function schemaTable(fields){
  const cols=[1900,1350,4150,1560];
  const header=['שדה','טיפוס','תיאור','מקור/הערה'];
  const rows=fields.map(f=>{
    const nameCell=[new Paragraph({bidirectional:true,alignment:AlignmentType.RIGHT,spacing:{after:20,before:20},children:[code(f[0],{size:18,bold:true,color:NAVY})]})];
    const typeCell=[new Paragraph({bidirectional:true,alignment:AlignmentType.RIGHT,spacing:{after:20,before:20},children:[code(f[1],{size:17,color:ACCENT})]})];
    return [nameCell,typeCell,f[2],f[3]];
  });
  return table(cols,header,rows);
}
function link(label,url){
  return new Paragraph({bidirectional:true,alignment:AlignmentType.RIGHT,spacing:{after:80,line:288},numbering:{reference:'b',level:0},
    children:[tR(label+' — '),new ExternalHyperlink({link:url,children:[new TextRun({text:url,font:'Consolas',size:17,color:'2E5496',underline:{}})]})]});
}

const c=[];
// ---- Title ----
c.push(new Paragraph({spacing:{before:1700},children:[new TextRun('')]}));
c.push(new Paragraph({bidirectional:true,alignment:AlignmentType.CENTER,spacing:{after:120},children:[new TextRun({text:'מסמך אפיון ותכנון',rightToLeft:true,font:FONT,size:30,bold:true,color:ACCENT})]}));
c.push(new Paragraph({bidirectional:true,alignment:AlignmentType.CENTER,spacing:{after:60},children:[new TextRun({text:'אטלס טמפורלי של האימפריה הרומית',rightToLeft:true,font:FONT,size:46,bold:true,color:NAVY})]}));
c.push(new Paragraph({bidirectional:true,alignment:AlignmentType.CENTER,spacing:{after:300},children:[new TextRun({text:'מפה דינמית לאורך כל תקופת רומא — ערים שצצות עם היווסדן',rightToLeft:true,font:FONT,size:26,color:BLUE})]}));
c.push(new Paragraph({bidirectional:true,alignment:AlignmentType.CENTER,spacing:{after:100},children:[new TextRun({text:'מפרט טכני, מודל נתונים טמפורלי ומקורות מידע',rightToLeft:true,font:FONT,size:22,italics:true,color:'555555'})]}));
c.push(new Paragraph({bidirectional:true,alignment:AlignmentType.CENTER,spacing:{before:1300,after:40},children:[new TextRun({text:'גרסה 1.0',rightToLeft:true,font:FONT,size:22,bold:true})]}));
c.push(new Paragraph({bidirectional:true,alignment:AlignmentType.CENTER,children:[new TextRun({text:'יולי 2026',rightToLeft:true,font:FONT,size:20,color:'555555'})]}));
c.push(new Paragraph({children:[new PageBreak()]}));

// ---- TOC ----
c.push(h1('תוכן עניינים'));
c.push(new TableOfContents('תוכן עניינים',{hyperlink:true,headingStyleRange:'1-3'}));
c.push(new Paragraph({children:[new PageBreak()]}));

// ---- 0 meta ----
c.push(h1('0. פרטי המסמך'));
c.push(table([2400,6560],['שדה','ערך'],[
  ['שם הפרויקט','Roman Empire Temporal Atlas (RE-Atlas · Time)'],
  ['סוג','מסמך אפיון ותכנון להרחבה טמפורלית'],
  ['מבוסס על','המפה האינטראקטיבית הקיימת (חתך זמן יחיד, ~200 לספירה)'],
  ['מטרה','מעבר ממפה של נקודת זמן אחת למפה דינמית על פני כל תקופת רומא'],
  ['קהל יעד','מנהל מוצר, אדריכל, צוות נתונים/ETL, מפתחי Frontend/Backend, עורך תוכן היסטורי'],
  ['סטטוס','טיוטה לאישור'],
]));
c.push(new Paragraph({children:[new PageBreak()]}));

// ---- 1 intro ----
c.push(h1('1. מבוא ותקציר מנהלים'));
c.push(h2('1.1 מטרת המסמך'));
c.push(p('מסמך זה מגדיר את כל מה שנדרש כדי להפוך את המפה הקיימת — שמתארת את האימפריה הרומית בנקודת זמן אחת (בקירוב 200 לספירה) — למפה טמפורלית דינמית המשתרעת על פני כל אורך תקופת רומא. באמצעות ציר זמן, המשתמש יוכל "להריץ" את ההיסטוריה ולראות כיצד ערים קמות (ומופיעות על המפה רק כשנוסדו), כיצד גבולות הפרובינציות משתנים, כיצד רשת הדרכים והלגיונות מתפתחת, וכיצד זרמי הסחר נעים לאורך הדורות. המסמך מפרט את החזון, מודל הנתונים הטמפורלי, מקורות המידע האקדמיים הנדרשים, האתגרים (בעיקר קנה-מידה של עשרות אלפי יישובים ושלמות תאריכים), והארכיטקטורה.'));

c.push(h2('1.2 מה משתנה לעומת הגרסה הקיימת'));
c.push(table([3100,2900,2960],['היבט','היום (סטטי)','היעד (טמפורלי)'],[
  ['ציר זמן','חתך יחיד ~200 לספירה','753 לפנה"ס – 476/1453 לספירה, עם סליידר'],
  ['יישובים','~90 ערים נבחרות','כל ערי/עיירות התקופה (עשרות אלפים) + כפרים אופציונלי'],
  ['הופעת ערים','כולן תמיד','כל עיר מופיעה רק מרגע היווסדה ונעלמת אם ננטשה/נחרבה'],
  ['גבולות','גבול יחיד משוער','גבולות משתנים לפי snapshots היסטוריים (AWMC)'],
  ['דרכים/לגיונות/סחר','מצב יחיד','מתפתחים לאורך זמן עם תוקף (valid-from/valid-to)'],
]));

c.push(h2('1.3 יעדים ומדדי הצלחה'));
c.push(bullet('דיוק טמפורלי: כל ישות מוצגת רק בטווח הזמן שבו התקיימה בפועל, מבוסס-מקור.'));
c.push(bullet('חוויית "נגן היסטוריה": הזזת ציר הזמן או הרצת אנימציה מעדכנת את כל השכבות בצורה חלקה.'));
c.push(bullet([tR('קנה מידה: תמיכה ב-'),code('≥ 30,000'),tR(' יישובים (Pleiades) עם ביצועים חלקים באמצעות clustering ו-vector tiles.')]));
c.push(bullet([tR('מדדים: זמן טעינה ראשוני '),code('< 4s'),tR(' ; עדכון מצב בשינוי שנה '),code('< 300ms'),tR(' ; כיסוי תיארוך '),code('≥ 80%'),tR(' מהמקומות (בהתאם לכיסוי Pleiades).')]));
c.push(new Paragraph({children:[new PageBreak()]}));

// ---- 2 time scope ----
c.push(h1('2. היקף הזמן וחלוקה לתקופות'));
c.push(p('טווח הזמן המוצע נע מייסוד רומא המסורתי (753 לפנה"ס) ועד נפילת האימפריה המערבית (476 לספירה), עם אפשרות הרחבה עד 1453 (נפילת קונסטנטינופול) עבור הרצף הרומי-ביזנטי. המערכת תשתמש באוצר-המילים של תקופות Pleiades ככלי לתיארוך גס כאשר שנה מדויקת אינה ידועה.'));
c.push(h2('2.1 תקופות מפתח (Snapshots) המומלצות כברירת מחדל'));
c.push(table([2600,2100,4260],['תקופה','שנה מייצגת','מאפיין'],[
  ['הרפובליקה המוקדמת','~350 לפנה"ס','רומא כמעצמה איטלקית'],
  ['שיא הרפובליקה','~60 לפנה"ס','ערב מלחמות האזרחים (snapshot AWMC 60 BCE)'],
  ['אוגוסטוס','14 לספירה','ייסוד הפרינקיפט'],
  ['שיא טראיאנוס','117 לספירה','ההיקף הטריטוריאלי המרבי (snapshot AWMC 117)'],
  ['השושלת הסוורית','200 לספירה','המצב במפה הקיימת (snapshot AWMC 200)'],
  ['רפורמת דיוקלטיאנוס','~300 לספירה','ריבוי פרובינציות (snapshot AWMC post-Diocletian)'],
  ['נפילת המערב','476 לספירה','התכווצות והתפרקות'],
]));
c.push(p([tR('אוצר-המילים של Pleiades מגדיר, בין היתר: '),code('Archaic'),tR(' (750–550 לפנה"ס), '),code('Hellenistic/Roman Republic'),tR(' (330–30 לפנה"ס), '),code('Roman'),tR(', ו-'),code('Late-Antique'),tR(' (300–640 לספירה). אלה משמשים לתיוג מקומות ללא שנת ייסוד מדויקת.')]));

c.push(h2('2.2 מוקש טכני: אין "שנת אפס" (הערה למפתחים)'));
c.push(p([tR('אזהרה קריטית לצוות הפיתוח: בלוח השנה ההיסטורי '),tR('אין שנת אפס',{bold:true}),tR('. השנה שאחרי 1 לפנה"ס (1 BCE) היא 1 לספירה (1 CE). מכיוון ש-'),code('foundedYear'),tR(' משתמש במספרים שליליים לתקופה שלפני הספירה, ביצוע חשבון פשוט על ציר הזמן (למשל הפרש שנים או מיפוי הסליידר) ייצור '),tR('באג של שנה אחת בדיוק סביב תקופת אוגוסטוס',{bold:true}),tR('. יש להגדיר פונקציית מעבר שמדלגת על אפס, ולהשתמש בה בכל חישוב זמן ובמיפוי הסליידר:')]));
c.push(new Paragraph({shading:{type:ShadingType.CLEAR,fill:'F5F2EA',color:'auto'},spacing:{before:60,line:240},indent:{left:120,right:120},children:[new TextRun({text:'// המרת שנה היסטורית (שלילי=לפנה"ס, ללא 0) לאינדקס רציף',font:'Consolas',size:18,color:'557755'})]}));
c.push(new Paragraph({shading:{type:ShadingType.CLEAR,fill:'F5F2EA',color:'auto'},spacing:{line:240},indent:{left:120,right:120},children:[new TextRun({text:'function histToIndex(y){ return y < 0 ? y + 1 : y; } // 1 BCE(-1)->0, 1 CE(1)->1',font:'Consolas',size:18,color:'333333'})]}));
c.push(new Paragraph({shading:{type:ShadingType.CLEAR,fill:'F5F2EA',color:'auto'},spacing:{after:140,line:240},indent:{left:120,right:120},children:[new TextRun({text:'function yearsBetween(a,b){ return histToIndex(b) - histToIndex(a); }',font:'Consolas',size:18,color:'333333'})]}));
c.push(p([tR('כל השוואות התוקף ('),code('validFrom ≤ Y ≤ validTo'),tR('), חישובי משך, ומיקום הסמן על ציר הזמן — יעברו דרך '),code('histToIndex'),tR(', כדי שהמעבר 1 לפנה"ס → 1 לספירה יהיה רציף וללא קפיצה או כפילות.')]));
c.push(new Paragraph({children:[new PageBreak()]}));

// ---- 3 data sources ----
c.push(h1('3. מקורות המידע (Ground Truth)'));
c.push(p('הצלחת המפה הטמפורלית תלויה בשילוב מספר מאגרים אקדמיים משלימים. הטבלה הבאה מרכזת את המקורות, מה כל אחד תורם, פורמט הנתונים, הרישוי והיקף.'));
c.push(table([1500,2650,2450,2360],['מקור','תורם','פורמט / היקף','רישוי'],[
  ['Pleiades','גזטיר המקומות המרכזי: עשרות אלפי ערים/אתרים עם תיארוך תקופתי ומזהה-על (Pleiades ID)','JSON מלא + CSV · ~40k מקומות, ~80% מתוארכים','CC BY'],
  ['Vici.org','עיירות, כפרים ואתרים ארכיאולוגיים ("אולי כפרים")','RDF/XML · ~20k מיקומים','CC BY-SA / CC0 מטא'],
  ['DARMC / MAPS (Harvard)','יישובים, דרכים (7,154 מקטעי דרך ב-117), מתקנים צבאיים, מכרות, וילות, אמות מים','Shapefile / XLSX · Harvard Dataverse','פתוח'],
  ['AWMC (Barrington)','גבולות והיקף האימפריה בחתכי זמן: 60 לפנה"ס, 117, 200, פוסט-דיוקלטיאנוס','GeoJSON · ODbL','ODbL'],
  ['ORBIS','עלויות וזמני מסע ברשת התחבורה (מודל סטטי, בסיס ~200 לספירה)','נתוני רשת','אקדמי'],
  ['OxREP','ספינות טרופות (Shipwrecks) ומכרות לאורך ציר הזמן — פרוקסי לנפח הסחר הימי ולפעילות הכלכלית לפי מאה','מסדי Excel/CSV · ~1,400 מכרות + אלפי טרופות','אקדמי'],
  ['Wikidata','שנות ייסוד מדויקות (inception, P571) היכן שידועות; קישור למזהי Pleiades/Vici','SPARQL / JSON','CC0'],
  ['Trismegistos','מקומות (בעיקר מצרים/פפירולוגיה) עם עדות ותיארוך','API / dumps','אקדמי'],
]));
c.push(p([tR('הבחנה חשובה: '),code('ORBIS'),tR(' הוא מודל '),tR('סטטי',{bold:true}),tR(' של עלויות וזמני תחבורה בשיא האימפריה — הוא אינו מתאר את השינוי בנפח הסחר לאורך מאות שנים. '),code('OxREP'),tR(' משלים בדיוק את החסר: נתוני הספינות הטרופות והמכרות שלו, המתוארכים לפי מאה, מאפשרים לגזור את '),tR('נפח',{bold:true}),tR(' הסחר הימי לאורך הזמן — כך שקווי הסחר הימיים יתעבו במאה ה-1–2 לספירה ויידקו ויֵחלשו במאה ה-3 עם קריסת האימפריה (ראה 6.5).')]));
c.push(h3('3.1 עקרון החיבור בין המקורות'));
c.push(p([tR('המפתח-על לאיחוד הוא '),code('Pleiades ID'),tR('. רוב המקורות (Vici, DARMC, Wikidata, Trismegistos) מקושרים אליו או ניתנים לקישור. תהליך ה-ETL יבצע התאמה לפי Pleiades ID כשקיים, ובגיבוי לפי קרבה גאוגרפית והתאמת שם מנורמל. שנת הייסוד תילקח לפי סדר עדיפות: Wikidata inception (מדויק) ← עדות ארכיאולוגית ← טווח תקופת Pleiades (גס).')]));
c.push(new Paragraph({children:[new PageBreak()]}));

// ---- 4 temporal data model ----
c.push(h1('4. מודל הנתונים הטמפורלי'));
c.push(p('כל ישות במערכת מקבלת "תוקף בזמן". העיקרון: כל אובייקט קיים בטווח [validFrom, validTo], וייתכנו לו מספר "מצבים" (states) לאורך חייו (שינוי שם, סטטוס, שיוך פרובינציאלי). התצוגה בשנה נתונה Y מציגה רק אובייקטים שבהם validFrom ≤ Y ≤ validTo.'));
c.push(h2('4.1 ישות Place (יישוב) — הרחבה טמפורלית'));
c.push(schemaTable([
  ['id','string','מזהה פנימי','פנימי'],
  ['pleiadesId','string','מזהה-על לקישור בין המקורות','Pleiades'],
  ['name','string','שם עכשווי לתצוגה','Pleiades'],
  ['nameHistory','object[]','שמות לאורך זמן: {name, from, to}','Pleiades/עריכה'],
  ['coordinates','number[2]','[lon, lat]','Pleiades/DARE'],
  ['foundedYear','integer?','שנת ייסוד (שלילי=לפנה"ס). null אם לא ידוע','Wikidata/ארכ׳'],
  ['foundedConfidence','enum','exact | approx | period','נגזר'],
  ['abandonedYear','integer?','שנת נטישה/חורבן אם ידועה','Wikidata/ארכ׳'],
  ['periodTags','string[]','תקופות Pleiades (archaic, roman, late-antique...)','Pleiades'],
  ['rank','enum','city | town | village | site | fort','Pleiades/Vici'],
  ['rankHistory','object[]','שינויי מעמד: {rank, from, to} (כפר→עיר)','עריכה'],
  ['provinceHistory','object[]','שיוך פרובינציאלי משתנה: {provinceId, from, to}','AWMC'],
  ['ethnicProfile','object[]','הרכב אתני-תרבותי משתנה בזמן: {groups:[{name,pct,lang}], from, to} (ראה 4.4)','אפיגרפיה/עריכה'],
  ['sourceIds','string[]','ייחוס למקורות','פנימי'],
]));
c.push(h2('4.2 ישות Province — גבולות משתנים בזמן'));
c.push(p('פרובינקיה אינה פוליגון יחיד אלא סדרת "גרסאות גבול" (borderVersions), אחת לכל snapshot היסטורי. בין snapshots ניתן להציג את הגבול הקרוב ביותר (step) או אינטרפולציה חזותית זהירה (ראה 6.3).'));
c.push(schemaTable([
  ['id','string','מזהה פרובינקיה','פנימי'],
  ['name','string','שם','AWMC/עריכה'],
  ['createdYear','integer','שנת הקמת הפרובינקיה','היסטורי'],
  ['dissolvedYear','integer?','שנת פירוק/פיצול','היסטורי'],
  ['borderVersions','object[]','{snapshotYear, polygon} — 60BCE/117/200/300...','AWMC'],
  ['econByEra','object','מדד עושר יחסי לכל תקופה','נגזר'],
  ['culturalDominance','object[]','תרבות שלטת לאורך זמן: {culture, from, to} (קלטי→גאלו-רומאי) (ראה 4.4)','אפיגרפיה/עריכה'],
]));
c.push(h2('4.3 ישויות Route / Legion / EconomicZone'));
c.push(bullet([code('Route'),tR(': שדות '),code('validFrom / validTo'),tR(' (דרך נסללה/ננטשה), ורשימת סחורות שיכולה להשתנות לפי תקופה.')]));
c.push(bullet([code('Legion'),tR(': היסטוריית פריסה '),code('deployments: [{baseId, from, to, legionName}]'),tR(' — לגיונות נעים בין בסיסים לאורך הדורות.')]));
c.push(bullet([code('EconomicZone'),tR(': '),code('activeFrom / activeTo'),tR(' — אזורי ייצור עולים ודועכים (למשל מכרות שמתמצים).')]));

c.push(h2('4.4 השכבה האתנית-תרבותית (Cultural Dominance) — מימד טמפורלי'));
c.push(p([tR('מעבר לשינויים הפיזיים (שמות, גבולות, מעמד), המימד האנושי-תרבותי הוא מה שהופך את המפה למרתקת. השדות '),code('Place.ethnicProfile'),tR(' ו-'),code('Province.culturalDominance'),tR(' מתעדים את ההרכב האתני והתרבות השלטת '),tR('לאורך זמן',{bold:true}),tR(', ומאפשרים למפה להציג ויזואלית תהליכי רומניזציה והלניזציה.')]));
c.push(bullet([code('ethnicProfile'),tR(' (רמת יישוב): עוגת הרכב אתני '),tR('שמשתנה עם הזמן',{italics:true}),tR(' — כל קבוצה עם אחוז ושפה, בתוקף לתקופה מסוימת.')]));
c.push(bullet([code('culturalDominance'),tR(' (רמת פרובינקיה): התרבות השלטת בכל תקופה, המשמשת לצביעה או לדוגמה (pattern) של הפרובינקיה על המפה.')]));
c.push(bullet([tR('דוגמה מערבית: '),tR('"הכתם הקלטי"',{bold:true}),tR(' בגאליה עובר בהדרגה ל-'),tR('"גאלו-רומאי"',{bold:true}),tR(' לאורך מאה-שנתיים (שינוי צבע/דוגמה עם תזוזת הסליידר).')]));
c.push(bullet([tR('דוגמה מזרחית: ערי המזרח (אנטיוכיה, אפסוס) '),tR('נשארות מתויגות תחת ',{}),code('Hellenistic/Greek'),tR(' לאורך כל התקופה — הרומניזציה כמעט אינה חודרת אליהן.')]));
c.push(p([tR('מקור: הצלבת עדות אפיגרפית ולשונית + Pleiades + ספרות מחקרית; כאשר אין נתון מפורש, נגזר לפי תקופה וברירת מחדל אזורית. שדה זה נועד לתצוגה איכותית (סיפור), ולא כנתון דמוגרפי מדויק — ויסומן ככזה.')]));
c.push(new Paragraph({children:[new PageBreak()]}));

// ---- 5 emergence logic ----
c.push(h1('5. לוגיקת "ערים צצות כשנבנו"'));
c.push(p('זהו הלב של הבקשה. הכלל: בשנה הנבחרת Y, יישוב מוצג רק אם התקיים אז.'));
c.push(num([tR('אם ידועה '),code('foundedYear'),tR(': הצג את היישוב כאשר '),code('Y ≥ foundedYear'),tR(' (ועד '),code('abandonedYear'),tR(' אם קיים).')]));
c.push(num([tR('אם ידוע רק '),code('periodTags'),tR(': הצג בטווח התקופות (למשל "Roman" → 30 לפנה"ס עד 300 לספירה), עם סימון חזותי של אי-ודאות (קו מקווקו / שקיפות).')]));
c.push(num([tR('אם אין תיארוך כלל: ברירת מחדל להצגה רק בחלון רחב (או הסתרה מאחורי מתג "כלול לא-מתוארכים").')]));
c.push(num('אנימציית הופעה: כאשר עיר "נכנסת" לחלון הזמן, הצג אפקט הופעה עדין (fade/pulse) כדי שהמשתמש יבחין בהיווסדות.'));
c.push(p([tR('הצגה של אי-ודאות היא עיקרון מנחה: מכיוון שלרבים מהמקומות אין שנת ייסוד מדויקת, המערכת חייבת להבחין חזותית בין "נוסד בשנה X (ודאי)" ל-"קיים בתקופה זו (משוער)". שדה '),code('foundedConfidence'),tR(' מניע את הסימון.')]));
c.push(new Paragraph({children:[new PageBreak()]}));

// ---- 6 UX ----
c.push(h1('6. חוויית משתמש וממשק זמן'));
c.push(h2('6.1 ציר הזמן (Time Slider)'));
c.push(bullet('סרגל זמן רציף בתחתית המסך: 753 לפנה"ס ← → 476 לספירה, עם סימוני תקופות מפתח.'));
c.push(bullet('גרירה מעדכנת את כל השכבות מיידית; תיבת קלט לשנה מדויקת.'));
c.push(bullet('כפתורי קפיצה ל-snapshots (אוגוסטוס, טראיאנוס 117, סוורוס 200, דיוקלטיאנוס).'));
c.push(h2('6.2 נגן אנימציה'));
c.push(bullet('Play/Pause, בקרת מהירות (שנים לשנייה), ולולאה. במהלך ההרצה ערים צצות, גבולות נעים, דרכים נסללות.'));
c.push(bullet('מונה שנה חי + תווית התקופה הנוכחית + סטטיסטיקה מתעדכנת (מס\' ערים פעילות, פרובינציות, לגיונות).'));
c.push(h2('6.3 גבולות משתנים ואינטרפולציה'));
c.push(p('בין שני snapshots (למשל 117 ל-200) יש שתי אפשרויות: (א) מצב מדרגה — הצג את הגבול של ה-snapshot האחרון עד שמגיעים לבא; (ב) אינטרפולציה חזותית (morph) — מומלץ להשתמש בזהירות ולסמן כ"משוער", שכן שינויי גבול היסטוריים לא היו ליניאריים.'));
c.push(h2('6.4 שכבות שמושפעות מהזמן'));
c.push(table([2400,6560],['שכבה','התנהגות בזמן'],[
  ['יישובים','הופעה בהיווסדות, היעלמות בנטישה, שינוי גודל/מעמד (כפר→עיר)'],
  ['פרובינציות','החלפת פוליגון לפי snapshot; צביעת עושר לפי התקופה'],
  ['דרכים','הופעת מקטעים לפי שנת סלילה (Via Appia 312 לפנה"ס ואילך...)'],
  ['לגיונות','מעבר בסיסים לאורך זמן; שינוי מספר הלגיונות הכולל'],
  ['סחר','שינוי היקפים ונתיבים לפי תקופה (ראה 6.5)'],
]));

c.push(h2('6.5 "עיוורון ימי" — ייצוג דינמי של הסחר הימי'));
c.push(p([tR('כדי להדגיש למשתמש את חשיבות הים לכלכלה הרומית, מקטעי הדרך הימיים יקבלו '),tR('ייצוג דינמי לאורך הזמן',{bold:true}),tR(', המקושר לנפח הסחר באותה מאה (נגזר מנתוני OxREP — ספינות טרופות לפי מאה כפרוקסי לפעילות ימית).')]));
c.push(bullet([tR('עובי קו ('),code('stroke-width'),tR(') ואנימציית זרימה ('),code('flow-animation'),tR(') של כל נתיב ימי — עוצמתם ומהירותם פרופורציונליות לנפח הסחר במאה הנוכחית.')]));
c.push(bullet([tR('שיא (מאה 1–2 לספירה): '),tR('"אוטוסטרדות כחולות"',{bold:true}),tR(' עבות וזורמות במהירות בין אלכסנדריה, רומא, קרתגו ואנטיוכיה.')]));
c.push(bullet([tR('משבר המאה ה-3: הנתיבים הימיים '),tR('"מתייבשים"',{bold:true}),tR(' — הקווים מדקדקים, הזרימה מאטה, וחלקם נעלמים כליל.')]));
c.push(bullet([tR('ניגוד חזותי מכוון: בו-זמנית, אייקוני הלגיונות '),tR('נודדים פנימה אל היבשה',{bold:true}),tR(' — מהחוף אל גבולות הריין, הדנובה והמזרח — כך שהמשתמש רואה את הים קורס בזמן שהצבא מתכנס אל היבשה.')]));
c.push(p([tR('אפקט זה הוא ליבת "הסיפור" של המפה הטמפורלית: הוא ממחיש כיצד הפריחה הימית של הפאקס רומאנה מתחלפת בהתכנסות צבאית-יבשתית של המשבר.')]));
c.push(new Paragraph({children:[new PageBreak()]}));

// ---- 7 scale/performance ----
c.push(h1('7. קנה מידה וביצועים'));
c.push(p('הוספת כל ערי/עיירות/כפרי התקופה משנה סדר גודל: מ-~90 אובייקטים ל-30,000–60,000. זהו האתגר ההנדסי המרכזי.'));
c.push(h2('7.1 אסטרטגיות'));
c.push(bullet([tR('Vector Tiles ('),code('PMTiles / MBTiles'),tR('): הגשת היישובים כאריחים וקטוריים במקום GeoJSON ענק — טעינה רק של מה שנראה.')]));
c.push(bullet('Clustering ורמות פירוט (LOD): בזום נמוך רק ערים גדולות; בהגדלה נחשפים עיירות וכפרים.'));
c.push(bullet('אינדוקס טמפורלי: סינון בצד השרת לפי שנה + bbox, כך שרק אובייקטים פעילים בשנה הנבחרת נשלחים.'));
c.push(bullet('דה-בּאונס (debounce) על גרירת הסליידר; טעינה מדורגת בזמן אנימציה.'));
c.push(h2('7.2 המלצות טכנולוגיה (לא מחייב)'));
c.push(bullet([code('MapLibre GL JS'),tR(' עם מקורות vector-tile; '),code('PMTiles'),tR(' לאחסון סטטי זול.')]));
c.push(bullet([code('PostgreSQL + PostGIS'),tR(' עם עמודות '),code('valid_from/valid_to'),tR(' ואינדקסים מרחביים+טמפורליים; או '),code('tippecanoe'),tR(' ליצירת האריחים.')]));
c.push(bullet('שכבת API: endpoint אחד עמוס — /features?year=Y&bbox=... — מחזיר רק ישויות פעילות.'));
c.push(new Paragraph({children:[new PageBreak()]}));

// ---- 8 architecture / ETL ----
c.push(h1('8. ארכיטקטורה ותהליך הנתונים (Pipeline)'));
c.push(h2('8.1 שלבי ה-ETL'));
c.push(num('Extract: הורדת Pleiades JSON, Vici RDF, DARMC shapefiles, AWMC snapshots, שאילתות Wikidata SPARQL לתאריכי ייסוד.'));
c.push(num([tR('Match: איחוד לפי '),code('pleiadesId'),tR('; גיבוי בהתאמה גאוגרפית + שם.')]));
c.push(num('Date: הקצאת foundedYear/abandonedYear + periodTags + confidence לכל מקום לפי סדר העדיפות (8.2).'));
c.push(num('Transform: נרמול סכמה, המרת יחידות, תרגום שמות תצוגה לעברית, חישוב rank והיסטוריית מעמד.'));
c.push(num([tR('Index & Tile: כתיבה ל-PostGIS + הפקת vector tiles ('),code('tippecanoe'),tR(') עם מאפייני '),code('from/to'),tR(' לכל feature.')]));
c.push(num('Validate: בדיקת שלמות תאריכים, טווחים הגיוניים, מפתחות זרים, כיסוי (ראה פרק 10).'));
c.push(h2('8.2 סדר עדיפות לתיארוך מקום'));
c.push(table([1200,3200,4560],['עדיפות','מקור','שימוש'],[
  ['1','Wikidata inception (P571)','שנת ייסוד מדויקת כשקיימת'],
  ['2','עדות ארכיאולוגית / ספרות','floruit או שכבת יסוד'],
  ['3','טווח תקופת Pleiades','תיארוך גס (period range) + confidence=period'],
  ['4','אין נתון','דגל "לא מתוארך"; הסתרה כברירת מחדל'],
]));
c.push(new Paragraph({children:[new PageBreak()]}));

// ---- 9 milestones ----
c.push(h1('9. אבני דרך ותכולת גרסאות'));
c.push(table([1500,2300,5160],['שלב','כותרת','תכולה'],[
  ['Phase 0','תשתית זמן','הוספת ציר זמן + שדות valid-from/to למודל הקיים; snapshots של AWMC לגבולות.'],
  ['Phase 1 (MVP)','גבולות ולגיונות בזמן','החלפת גבולות בין 4 snapshots; לגיונות ודרכים עם תוקף; נגן אנימציה בסיסי.'],
  ['Phase 2','כל ערי Pleiades','קליטת גזטיר Pleiades המלא + תיארוך; vector tiles; clustering; ערים צצות בהיווסדן.'],
  ['Phase 3','עיירות וכפרים','שילוב Vici.org ו-DARMC לרזולוציית עיירות/כפרים; חשיפה לפי זום.'],
  ['Phase 4','ליטוש','אינטרפולציית גבולות, אנימציות הופעה, סימון אי-ודאות, ביצועים ונגישות.'],
]));
c.push(h2('9.1 קריטריוני קבלה ל-MVP'));
c.push(bullet('הזזת הסליידר בין 60 לפנה"ס, 117, 200, 300 מחליפה את גבולות הפרובינציות בהתאם.'));
c.push(bullet('ערים מופיעות/נעלמות לפי foundedYear/abandonedYear.'));
c.push(bullet('נגן אנימציה מריץ את הזמן וכל השכבות מתעדכנות חלק.'));
c.push(new Paragraph({children:[new PageBreak()]}));

// ---- 10 risks ----
c.push(h1('10. סיכונים, הנחות ותלויות'));
c.push(h2('10.1 הנחות'));
c.push(bullet('הנתונים זמינים להורדה ושימוש בכפוף לרישיונם (Pleiades CC BY, AWMC ODbL, Vici CC BY-SA).'));
c.push(bullet('תיארוך גס (רמת תקופה) מקובל עבור רוב היישובים; דיוק שנה נדרש רק לערים מרכזיות.'));
c.push(h2('10.2 סיכונים'));
c.push(table([3200,1400,4360],['סיכון','חומרה','מיטיגציה'],[
  ['חוסר שלמות בתאריכי ייסוד','גבוהה','שימוש בטווחי תקופות + סימון אי-ודאות + מתג "כלול לא-מתוארכים".'],
  ['ביצועים עם עשרות אלפי נקודות','גבוהה','Vector tiles, clustering, סינון שרת לפי שנה+bbox.'],
  ['גבולות בין snapshots לא ליניאריים','בינונית','ברירת מחדל למצב מדרגה; אינטרפולציה מסומנת כמשוערת בלבד.'],
  ['אי-התאמת מזהים בין מקורות','בינונית','Pleiades ID כמפתח-על + סקירה ידנית של מקרי קצה.'],
  ['עומס עריכה (תרגום/תיאור)','בינונית','התמקדות בערים מרכזיות; שאר המקומות עם מטא-דאטה בלבד.'],
]));
c.push(h2('10.3 תלויות'));
c.push(bullet('זמינות והרשאות של Pleiades, AWMC, Vici, DARMC, Wikidata, Trismegistos.'));
c.push(bullet('מנוע מפות התומך vector tiles ואריחי בסיס מתאימים לעולם העתיק.'));
c.push(new Paragraph({children:[new PageBreak()]}));

// ---- 11 appendix sources ----
c.push(h1('11. נספח: מקורות אקדמיים וקישורים'));
c.push(p('המקורות הבאים משמשים כ-Ground Truth למפה הטמפורלית:'));
c.push(link('Pleiades — גזטיר המקומות העתיקים (הורדות)','https://pleiades.stoa.org/downloads'));
c.push(link('Pleiades — אוצר מילים של תקופות זמן','https://pleiades.stoa.org/vocabularies/time-periods'));
c.push(link('Vici.org — אטלס ארכיאולוגי של העת העתיקה','https://vici.org'));
c.push(link('DARMC / Mapping Past Societies (Harvard)','https://darmc.harvard.edu'));
c.push(link('DARMC — Harvard Dataverse (הורדות)','https://dataverse.harvard.edu/dataverse/darmc'));
c.push(link('AWMC — נתוני גבולות בחתכי זמן (Barrington)','https://github.com/AWMC/geodata'));
c.push(link('ORBIS — מודל רשת התחבורה','https://orbis.stanford.edu'));
c.push(link('OxREP — פרויקט הכלכלה הרומית (ספינות טרופות ומכרות)','https://oxrep.web.ox.ac.uk'));
c.push(link('Wikidata — מאפיין inception (P571)','https://www.wikidata.org/wiki/Property:P571'));
c.push(link('Trismegistos — מקומות העת העתיקה','https://www.trismegistos.org/geo/'));
c.push(spacer(200));
c.push(p([tR('— סוף המסמך —')],{align:AlignmentType.CENTER,color:'888888',italics:true}));

// ---- assemble ----
const doc=new Document({
  creator:'RE-Atlas',title:'אפיון אטלס טמפורלי — האימפריה הרומית',
  styles:{default:{document:{run:{font:FONT,size:22}}}},
  numbering:{config:[
    {reference:'b',levels:[
      {level:0,format:LevelFormat.BULLET,text:'•',alignment:AlignmentType.LEFT,style:{run:{color:ACCENT},paragraph:{indent:{right:360,hanging:220}}}},
      {level:1,format:LevelFormat.BULLET,text:'◦',alignment:AlignmentType.LEFT,style:{paragraph:{indent:{right:720,hanging:220}}}},
    ]},
    {reference:'n',levels:[
      {level:0,format:LevelFormat.DECIMAL,text:'%1.',alignment:AlignmentType.RIGHT,style:{run:{bold:true,color:NAVY},paragraph:{indent:{right:360,hanging:260}}}},
    ]},
  ]},
  sections:[{
    properties:{page:{size:{width:11906,height:16838},margin:{top:1200,bottom:1200,left:1100,right:1100}}},
    headers:{default:new Header({children:[new Paragraph({bidirectional:true,alignment:AlignmentType.LEFT,border:{bottom:{style:BorderStyle.SINGLE,size:4,color:'CCCCCC',space:4}},children:[new TextRun({text:'RE-Atlas · אפיון טמפורלי v1.0',rightToLeft:true,font:FONT,size:16,color:'888888'})]})]})},
    footers:{default:new Footer({children:[new Paragraph({alignment:AlignmentType.CENTER,children:[new TextRun({text:'עמוד ',rightToLeft:true,font:FONT,size:16,color:'888888'}),new TextRun({children:[PageNumber.CURRENT],font:FONT,size:16,color:'888888'})]})]})},
    children:c,
  }],
});
Packer.toBuffer(doc).then(buf=>{fs.writeFileSync(process.argv[2]||'out.docx',buf);console.log('WROTE',buf.length,'bytes');});
