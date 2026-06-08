# ทดสอบ Thai Document Skill

โฟลเดอร์นี้มีสคริปต์สร้างเอกสารทดสอบ 3 หน้า เพื่อยืนยันการทำงานของ skill ทั้ง 4 สคริปต์

## โครงสร้าง

```
tests/
├── create.js          ← สร้าง test-raw.docx ด้วย docx-js (Node.js)
├── package.json       ← dependency: docx
└── README.md          ← ไฟล์นี้
```

## ความต้องการ

```bash
node --version   # >= 16
python --version # >= 3.8
npm install      # ติดตั้ง docx library
pip install python-docx pythainlp
```

## วิธีรัน

```bash
cd tests
npm install

# 1. สร้างเอกสารดิบ
node create.js

# 2. ผ่านสคริปต์ทั้ง 4 ตามลำดับ
python ../scripts/page_margin.py  test-raw.docx --preset a4-thai      -o s1.docx
python ../scripts/table_fit.py    s1.docx       --mode auto            -o s2.docx
python ../scripts/apply_indent.py s2.docx                             -o s3.docx
python ../scripts/thai_wrap.py    s3.docx                             -o test-final.docx
```

## เนื้อหาเอกสาร 3 หน้า

| หน้า | หัวข้อ | สิ่งที่ทดสอบ |
|------|--------|--------------|
| 1 | บทนำ / วัตถุประสงค์ / ขอบเขต | ขอบกระดาษ a4-thai, ย่อหน้า 1.27 ซม., ZWSP ตัดคำ |
| 2 | ตารางกว้างเกิน (14,000 DXA) | table_fit ย่อให้พอดีหน้า (≈ 8,504 DXA) |
| 3 | ตารางขั้นตอน + สรุป | ตารางปกติ, ข้อความ LEFT+ZWSP เต็มบรรทัด |

## ผลลัพธ์ที่ควรได้

```
OK: set margins on 1 section(s): top=2.54 bottom=2.54 left=3.0 right=2.0 cm
OK: checked 2 table(s) | scaled: 1 | autofit: 0 | ok: 1
    text area: 453.6 pt (16.00 cm)
✓  ตั้งย่อหน้า 1.27 ซม. → 15 ย่อหน้า
✓  ตัดคำแล้ว 77 run (ข้าม 0 run ใน justified para)
```

## หมายเหตุสำคัญ: LEFT vs JUSTIFY ใน .docx

เอกสารนี้ใช้ **LEFT alignment** สำหรับ body text ทั้งหมด ไม่ใช่ JUSTIFY

| Alignment | พฤติกรรม | เหมาะกับ |
|-----------|----------|----------|
| `LEFT` + ZWSP | ตัดบรรทัดที่ขอบเขตคำ ขอบขวาหยัก อ่านสะดวก | ✅ .docx |
| `JUSTIFY` + ZWSP | Word กระจาย space ที่ทุก ZWSP → ช่องว่างถ่าง | ❌ .docx |
| `JUSTIFY` + CSS `inter-character` | กระจาย space ระหว่างอักขระ ดูเรียบร้อย | ✅ HTML/PDF |
