const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
  Header, Footer, AlignmentType, HeadingLevel, BorderStyle, WidthType,
  ShadingType, VerticalAlign, PageNumber, PageBreak, LevelFormat
} = require("docx");
const fs = require("fs");

// ========================= ข้อมูลเนื้อหา =========================

const border = { style: BorderStyle.SINGLE, size: 1, color: "AAAAAA" };
const borders = { top: border, bottom: border, left: border, right: border };

function cell(text, widthDxa, { bold = false, shade = null, align = AlignmentType.LEFT } = {}) {
  return new TableCell({
    borders,
    width: { size: widthDxa, type: WidthType.DXA },
    shading: shade ? { fill: shade, type: ShadingType.CLEAR } : undefined,
    margins: { top: 80, bottom: 80, left: 120, right: 120 },
    verticalAlign: VerticalAlign.CENTER,
    children: [new Paragraph({
      alignment: align,
      children: [new TextRun({ text, bold, font: "TH Sarabun New", size: 28 })]
    })]
  });
}

// ---- ตารางหน้า 2: ข้อมูลที่กว้างเกินกระดาษ (ทดสอบ table_fit) ----
const wideTable = new Table({
  width: { size: 14000, type: WidthType.DXA },   // จงใจให้กว้างเกิน A4
  columnWidths: [2800, 2800, 2800, 2800, 2800],
  rows: [
    new TableRow({ children: [
      cell("ลำดับ",      2800, { bold: true, shade: "D6E4F0", align: AlignmentType.CENTER }),
      cell("ชื่อวิชา",    2800, { bold: true, shade: "D6E4F0" }),
      cell("หน่วยกิต", 2800, { bold: true, shade: "D6E4F0", align: AlignmentType.CENTER }),
      cell("ผู้สอน",     2800, { bold: true, shade: "D6E4F0" }),
      cell("คะแนนเต็ม", 2800, { bold: true, shade: "D6E4F0", align: AlignmentType.CENTER }),
    ]}),
    ...([
      ["1", "การออกแบบและจัดการเครือข่ายในองค์กร", "3", "อ.สมศักดิ์ ใจดี", "100"],
      ["2", "ระบบปฏิบัติการเครือข่าย", "3", "อ.วิภาวดี สุขใจ", "100"],
      ["3", "การเขียนโปรแกรมบนเครือข่าย", "3", "อ.ประสิทธิ์ เก่งมาก", "100"],
      ["4", "ความมั่นคงของระบบสารสนเทศ", "3", "อ.ณัฐพงศ์ ดีงาม", "100"],
      ["5", "โครงการพัฒนาระบบสารสนเทศองค์กร", "6", "อ.สมหญิง รักดี", "100"],
    ]).map(([n, name, cr, teacher, score]) =>
      new TableRow({ children: [
        cell(n,       2800, { align: AlignmentType.CENTER }),
        cell(name,    2800),
        cell(cr,      2800, { align: AlignmentType.CENTER }),
        cell(teacher, 2800),
        cell(score,   2800, { align: AlignmentType.CENTER }),
      ]})
    )
  ]
});

// ---- ตารางหน้า 3: ปกติ (พอดีหน้า) ----
const normalTable = new Table({
  width: { size: 8504, type: WidthType.DXA },
  columnWidths: [1200, 5504, 1800],
  rows: [
    new TableRow({ children: [
      cell("ขั้นตอน", 1200, { bold: true, shade: "E8F5E9", align: AlignmentType.CENTER }),
      cell("รายละเอียด", 5504, { bold: true, shade: "E8F5E9" }),
      cell("สคริปต์", 1800, { bold: true, shade: "E8F5E9", align: AlignmentType.CENTER }),
    ]}),
    ...([
      ["1", "ตั้งขอบกระดาษให้ถูกต้องตามมาตรฐาน ใช้ preset a4-thai สำหรับเอกสารราชการไทย", "page_margin.py"],
      ["2", "ป้องกันตารางล้นขอบกระดาษ ย่อสัดส่วนหรือตั้ง AutoFit อัตโนมัติ", "table_fit.py"],
      ["3", "ตั้งย่อหน้าบรรทัดแรกตามมาตรฐานสากล APA/MLA 1.27 ซม.", "apply_indent.py"],
      ["4", "แทรกจุดตัดคำ (ZWSP) กันบรรทัดตัดกลางคำภาษาไทย", "thai_wrap.py"],
    ]).map(([n, detail, script]) =>
      new TableRow({ children: [
        cell(n,      1200, { align: AlignmentType.CENTER }),
        cell(detail, 5504),
        cell(script, 1800, { align: AlignmentType.CENTER }),
      ]})
    )
  ]
});

// ========================= paragraph helpers =========================

const T = (text, opts = {}) => new TextRun({ text, font: "TH Sarabun New", size: 28, ...opts });

function P(children, opts = {}) {
  return new Paragraph({
    children: Array.isArray(children) ? children : [children],
    spacing: { after: 160 },
    ...opts
  });
}

function H(level, text) {
  return new Paragraph({
    heading: level,
    spacing: { before: 240, after: 120 },
    children: [new TextRun({ text, bold: true, font: "TH Sarabun New", size: level === HeadingLevel.HEADING_1 ? 40 : 32 })]
  });
}

function indent(text) {
  // LEFT alignment สำหรับ .docx ภาษาไทย:
  //   - JUSTIFY ใน Word กระจาย space ระหว่างคำ → ช่องว่างถ่างเมื่อบรรทัดไม่เต็ม
  //   - LEFT + ZWSP (จาก thai_wrap.py) = ตัดที่ขอบเขตคำพอดี สวยงามกว่า
  //   - JUSTIFY เหมาะกับ HTML/PDF ที่ใช้ CSS text-justify: inter-character เท่านั้น
  return new Paragraph({
    children: [T(text)],
    indent: { firstLine: 720 },
    alignment: AlignmentType.LEFT,
    spacing: { after: 200 }
  });
}

// ========================= สร้างเอกสาร =========================

const doc = new Document({
  styles: {
    default: {
      document: { run: { font: "TH Sarabun New", size: 28 } }
    },
    paragraphStyles: [
      {
        id: "Heading1", name: "Heading 1", basedOn: "Normal", next: "Normal", quickFormat: true,
        run: { size: 40, bold: true, font: "TH Sarabun New", color: "1F4E79" },
        paragraph: { spacing: { before: 240, after: 120 }, outlineLevel: 0 }
      },
      {
        id: "Heading2", name: "Heading 2", basedOn: "Normal", next: "Normal", quickFormat: true,
        run: { size: 32, bold: true, font: "TH Sarabun New", color: "2E75B6" },
        paragraph: { spacing: { before: 180, after: 80 }, outlineLevel: 1 }
      }
    ]
  },
  sections: [{
    properties: {
      page: {
        size: { width: 11906, height: 16838 },  // A4
        margin: { top: 1440, bottom: 1440, left: 1701, right: 1134 }  // บน/ล่าง 1", ซ้าย 3ซม., ขวา 2ซม.
      }
    },
    headers: {
      default: new Header({ children: [new Paragraph({
        alignment: AlignmentType.RIGHT,
        border: { bottom: { style: BorderStyle.SINGLE, size: 4, color: "2E75B6", space: 4 } },
        children: [T("เอกสารทดสอบ Thai Document Skill  |  DocumentThaiSkill", { color: "555555", size: 22 })]
      })] })
    },
    footers: {
      default: new Footer({ children: [new Paragraph({
        alignment: AlignmentType.CENTER,
        border: { top: { style: BorderStyle.SINGLE, size: 4, color: "AAAAAA", space: 4 } },
        children: [
          T("หน้า ", { size: 22, color: "555555" }),
          new TextRun({ children: [PageNumber.CURRENT], font: "TH Sarabun New", size: 22, color: "555555" }),
          T(" / ", { size: 22, color: "555555" }),
          new TextRun({ children: [PageNumber.TOTAL_PAGES], font: "TH Sarabun New", size: 22, color: "555555" }),
        ]
      })] })
    },
    children: [

      // ══════════════════════════════════════════
      //  หน้า 1 — ปกและบทนำ
      // ══════════════════════════════════════════

      new Paragraph({
        alignment: AlignmentType.CENTER,
        spacing: { before: 480, after: 240 },
        children: [new TextRun({ text: "รายงานทดสอบระบบจัดเอกสารภาษาไทย", bold: true, font: "TH Sarabun New", size: 52, color: "1F4E79" })]
      }),
      new Paragraph({
        alignment: AlignmentType.CENTER,
        spacing: { after: 80 },
        children: [T("DocumentThaiSkill — Thai Document Formatting Skill", { size: 30, color: "2E75B6" })]
      }),
      new Paragraph({
        alignment: AlignmentType.CENTER,
        spacing: { after: 480 },
        children: [T("วันที่ 8 มิถุนายน 2569  |  วิทยาลัยเทคนิคอุบลราชธานี", { size: 24, color: "777777" })]
      }),

      H(HeadingLevel.HEADING_1, "1. บทนำ"),

      indent("เอกสารนี้จัดทำขึ้นเพื่อทดสอบการทำงานของ DocumentThaiSkill ซึ่งเป็นชุดสคริปต์สำหรับจัดรูปแบบเอกสารภาษาไทยให้ถูกต้องตามมาตรฐาน ครอบคลุมการตัดคำ การกั้นหน้าหลัง การป้องกันตารางล้นขอบ และการตั้งย่อหน้าบรรทัดแรก"),

      indent("ปัญหาที่พบบ่อยในเอกสารภาษาไทยได้แก่ การตัดบรรทัดกลางคำ ตารางที่กว้างกว่ากระดาษ ขอบกระดาษที่ไม่ได้มาตรฐาน และการย่อหน้าที่ไม่สม่ำเสมอ สคริปต์ในชุดนี้แก้ปัญหาทั้งหมดนั้นได้ในขั้นตอนเดียว"),

      H(HeadingLevel.HEADING_2, "1.1 วัตถุประสงค์"),

      P([T("การทดสอบนี้มีวัตถุประสงค์ดังนี้")]),
      P([T("  ก) ทดสอบการตั้งขอบกระดาษด้วย page_margin.py ให้ได้มาตรฐาน a4-thai")]),
      P([T("  ข) ทดสอบการป้องกันตารางล้นขอบด้วย table_fit.py")]),
      P([T("  ค) ทดสอบการตั้งย่อหน้าสากลด้วย apply_indent.py")]),
      P([T("  ง) ทดสอบการตัดคำภาษาไทยด้วย thai_wrap.py")]),

      H(HeadingLevel.HEADING_2, "1.2 ขอบเขต"),

      indent("การทดสอบครอบคลุมเอกสาร .docx ที่มีเนื้อหาภาษาไทย ประกอบด้วยหัวข้อ ย่อหน้า ตาราง และการจัดรูปแบบหัวกระดาษและท้ายกระดาษ รวมทั้งการแสดงเลขหน้าอัตโนมัติ"),

      // ══════════════════════════════════════════
      //  หน้า 2 — ตารางกว้างเกิน (ทดสอบ table_fit)
      // ══════════════════════════════════════════

      new Paragraph({ children: [new PageBreak()] }),

      H(HeadingLevel.HEADING_1, "2. ทดสอบตารางกว้างเกิน (table_fit.py)"),

      indent("ตารางด้านล่างนี้มีความกว้างรวม 14,000 DXA ซึ่งกว้างกว่าพื้นที่ข้อความของกระดาษ A4 (ประมาณ 8,504 DXA) เมื่อรันผ่าน table_fit.py ด้วยโหมด auto ตารางจะถูกย่อสัดส่วนโดยอัตโนมัติ"),

      P([T("ตารางที่ 1: รายวิชาในหลักสูตร ปวส. เทคโนโลยีเครือข่าย (ก่อนผ่าน table_fit)", { bold: true, color: "C00000" })], { alignment: AlignmentType.CENTER }),

      wideTable,

      P([T("")]),

      indent("หลังจากรัน table_fit.py --mode auto ตารางข้างต้นจะถูกย่อสัดส่วนคอลัมน์ทุกคอลัมน์ให้ผลรวมความกว้างไม่เกินพื้นที่ข้อความ ขณะที่ยังรักษาอัตราส่วนระหว่างคอลัมน์ไว้เท่าเดิม"),

      H(HeadingLevel.HEADING_2, "2.1 คำสั่งที่ใช้"),

      P([T("python3 scripts/table_fit.py test.docx --mode auto -o test_fitted.docx", { font: "Courier New", size: 24, color: "1F4E79" })]),

      // ══════════════════════════════════════════
      //  หน้า 3 — สรุปขั้นตอนและผลลัพธ์
      // ══════════════════════════════════════════

      new Paragraph({ children: [new PageBreak()] }),

      H(HeadingLevel.HEADING_1, "3. สรุปขั้นตอนการจัดเอกสาร"),

      indent("ลำดับการทำงานที่แนะนำสำหรับเอกสารงานวิจัยและรายงานวิชาการภาษาไทย ควรดำเนินการตามลำดับต่อไปนี้เพื่อให้ได้ผลลัพธ์ที่ดีที่สุด"),

      P([T("ตารางที่ 2: ลำดับการใช้งานสคริปต์", { bold: true })], { alignment: AlignmentType.CENTER }),

      normalTable,

      P([T("")]),

      H(HeadingLevel.HEADING_2, "3.1 ตัวอย่างเนื้อหาที่ผ่านการตัดคำแล้ว"),

      indent("การออกแบบและการจัดการเครือข่ายในองค์กรเป็นวิชาที่สำคัญในหลักสูตรประกาศนียบัตรวิชาชีพชั้นสูง สาขาเทคโนโลยีสารสนเทศ ผู้เรียนจะได้ศึกษาหลักการออกแบบโทโพโลยีเครือข่าย การกำหนดที่อยู่ไอพี การติดตั้งและกำหนดค่าอุปกรณ์เครือข่าย รวมถึงการแก้ปัญหาที่เกิดขึ้นในระบบเครือข่ายจริง"),

      indent("เนื้อหาที่ผ่านการแทรก Zero-Width Space (ZWSP) โดย thai_wrap.py จะถูกตัดบรรทัดที่ขอบเขตคำเสมอ ไม่ว่าข้อความจะถูกแสดงในขนาดฟอนต์ใดหรือความกว้างหน้าเท่าใดก็ตาม ทำให้เอกสารดูเรียบร้อยและอ่านง่ายในทุกสภาพแวดล้อม"),

      H(HeadingLevel.HEADING_2, "3.2 สรุป"),

      indent("DocumentThaiSkill ช่วยให้การจัดรูปแบบเอกสารภาษาไทยเป็นไปโดยอัตโนมัติ ลดเวลาที่ต้องเสียไปกับการปรับแต่งด้วยมือ และทำให้เอกสารได้มาตรฐานสากลทั้งในด้านขอบกระดาษ ย่อหน้า และการตัดคำ เหมาะสำหรับการนำไปใช้ในงานราชการ รายงานวิชาการ และวิทยานิพนธ์"),

      new Paragraph({
        alignment: AlignmentType.CENTER,
        spacing: { before: 400 },
        children: [T("— สิ้นสุดเอกสารทดสอบ —", { color: "777777", italics: true })]
      })
    ]
  }]
});

Packer.toBuffer(doc).then(buf => {
  fs.writeFileSync("test-raw.docx", buf);
  console.log("✓ สร้าง test-raw.docx เรียบร้อย");
});
