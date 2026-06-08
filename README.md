# DocumentThaiSkill — thai-doc-wrap

Skill สำหรับ **จัดเอกสารภาษาไทยให้สวยงาม**: ตัดคำไม่ให้บรรทัดขาดกลางคำ และตั้งมาตรฐานการย่อหน้า (first-line indent) แบบสากลตามงานวิจัย รองรับไฟล์ `.txt`, `.html`, `.docx` และงานพิมพ์/PDF

> A Claude Skill that makes Thai documents wrap beautifully (no mid-word line breaks) and applies an international research-standard first-line indent. Works on `.txt`, `.html`, `.docx`, and print/PDF.

## ปัญหาที่แก้

ภาษาไทยเขียนติดกันไม่เว้นวรรคระหว่างคำ โปรแกรมจัดหน้าจึงตัดบรรทัด "กลางคำ" ทำให้อ่านยาก Skill นี้ใช้พจนานุกรมตัดคำ (PyThaiNLP) หาขอบเขตคำ แล้วแทรก **Zero-Width Space (U+200B)** ที่มองไม่เห็น เพื่อให้บรรทัดตัดที่ท้ายคำเสมอ — ได้ผลเหมือนกันทุกที่ รวมถึงตอนส่งออก PDF/HTML หรือเปิดในเครื่องที่ไม่มีชุดภาษาไทย

## โครงสร้าง

```
thai-doc-wrap/
├── SKILL.md                              # คำสั่ง/คู่มือของ skill
├── scripts/
│   ├── thai_wrap.py                      # ตัดคำ + แทรกจุดตัดบรรทัด (txt/html/docx)
│   └── apply_indent.py                   # ตั้งย่อหน้าบรรทัดแรกมาตรฐานสากล (docx)
├── assets/
│   └── thai-print.css                    # แม่แบบ CSS งานพิมพ์ไทย
└── references/
    └── research-indent-standard.md       # มาตรฐานย่อหน้า/แท็บ สากลตามงานวิจัย
```

`thai-doc-wrap.skill` ที่อยู่ใน repo คือไฟล์แพ็กเกจพร้อมติดตั้งใน Claude (Settings → Capabilities/Skills)

## ติดตั้ง dependencies

```bash
pip install pythainlp --break-system-packages -q   # จำเป็น (ตัดคำ)
# beautifulsoup4 (HTML) และ python-docx (DOCX) มักมีอยู่แล้ว
```

## การใช้งาน

ตัดคำ (กันบรรทัดตัดกลางคำ):
```bash
echo "การออกแบบและการจัดการเครือข่ายในองค์กร" | python3 thai-doc-wrap/scripts/thai_wrap.py
python3 thai-doc-wrap/scripts/thai_wrap.py report.docx -o report_wrapped.docx
python3 thai-doc-wrap/scripts/thai_wrap.py page.html  -o page_wrapped.html
```

ตั้งย่อหน้าบรรทัดแรกมาตรฐานสากล (APA/MLA/Chicago = 1.27 ซม. / 0.5"):
```bash
python3 thai-doc-wrap/scripts/apply_indent.py research.docx -o research_final.docx
python3 thai-doc-wrap/scripts/apply_indent.py thesis.docx --cm 1.0 -o thesis_final.docx   # วิทยานิพนธ์ มจพ.
```

ลำดับแนะนำในงานวิจัย: สร้างเนื้อหา → `apply_indent.py` → `thai_wrap.py` (ตัดคำเป็นขั้นสุดท้าย)

## มาตรฐานย่อหน้า

| รูปแบบ | ย่อหน้าบรรทัดแรก | นิ้ว | DXA |
|---|---|---|---|
| สากล (APA 7th / MLA 9th / Chicago) | 1.27 ซม. | 0.5" | 720 |
| วิทยานิพนธ์ มจพ. (KMUTNB) | 1.0 ซม. | 0.39" | 567 |

## License

MIT — ใช้ PyThaiNLP (Apache-2.0)
