"""
page_margin.py — ตั้งขอบกระดาษ (page margins) ให้ไฟล์ .docx

การใช้งาน:
  python3 scripts/page_margin.py input.docx -o output.docx
  python3 scripts/page_margin.py input.docx --preset a4-thai
  python3 scripts/page_margin.py input.docx --top 3 --bottom 2.5 --left 3 --right 2.5

หน่วยทุก argument คือเซนติเมตร (1 cm = 567 EMU × 100 / 2.54 ≈ 567 DXA)
"""

import argparse
import sys
from pathlib import Path

try:
    from docx import Document
    from docx.shared import Cm
    from docx.oxml.ns import qn
except ImportError:
    sys.exit("ต้องติดตั้ง python-docx ก่อน:  pip install python-docx")


# Preset margin sets (top, bottom, left, right) หน่วย cm
PRESETS = {
    "a4-thai": (2.54, 2.54, 3.0, 2.0),      # ราชการไทย: บน/ล่าง 1", ซ้าย 3 ซม., ขวา 2 ซม.
    "a4-standard": (2.54, 2.54, 2.54, 2.54), # สี่ด้านเท่ากัน 1"
    "thesis-mku": (3.0, 2.5, 4.0, 2.5),      # วิทยานิพนธ์ มจพ. (จากคู่มือ 2558)
    "report": (2.5, 2.5, 3.0, 2.5),          # รายงานทั่วไป
}


def set_margins(doc_path: Path, out_path: Path,
                top: float, bottom: float, left: float, right: float) -> int:
    doc = Document(str(doc_path))
    count = 0
    for section in doc.sections:
        section.top_margin = Cm(top)
        section.bottom_margin = Cm(bottom)
        section.left_margin = Cm(left)
        section.right_margin = Cm(right)
        count += 1
    doc.save(str(out_path))
    return count


def main():
    ap = argparse.ArgumentParser(description="ตั้งขอบกระดาษ .docx")
    ap.add_argument("input", help="ไฟล์ .docx ต้นทาง")
    ap.add_argument("-o", "--output", help="ไฟล์ผลลัพธ์ (ไม่ระบุ = เพิ่ม .margin.docx)")
    ap.add_argument("--preset", choices=list(PRESETS), default=None,
                    help="ชุดขอบสำเร็จรูป: " + ", ".join(PRESETS))
    ap.add_argument("--top",    type=float, help="บน (ซม.)")
    ap.add_argument("--bottom", type=float, help="ล่าง (ซม.)")
    ap.add_argument("--left",   type=float, help="ซ้าย (ซม.)")
    ap.add_argument("--right",  type=float, help="ขวา (ซม.)")
    args = ap.parse_args()

    src = Path(args.input)
    if not src.exists():
        sys.exit(f"ไม่พบไฟล์: {src}")

    # กำหนด margin
    if args.preset:
        top, bottom, left, right = PRESETS[args.preset]
    else:
        top    = args.top    if args.top    is not None else 2.54
        bottom = args.bottom if args.bottom is not None else 2.54
        left   = args.left   if args.left   is not None else 3.0
        right  = args.right  if args.right  is not None else 2.0

    # บน/ล่าง/ซ้าย/ขวา จากอาร์กิวเมนต์แต่ละตัวแทนค่าจาก preset
    if args.preset:
        if args.top    is not None: top    = args.top
        if args.bottom is not None: bottom = args.bottom
        if args.left   is not None: left   = args.left
        if args.right  is not None: right  = args.right

    dst = Path(args.output) if args.output else src.with_suffix(".margin.docx")

    sections = set_margins(src, dst, top, bottom, left, right)
    print(f"✓ ตั้งขอบกระดาษ {sections} section: บน={top} ล่าง={bottom} ซ้าย={left} ขวา={right} ซม.")
    print(f"  บันทึกแล้ว → {dst}")


if __name__ == "__main__":
    main()
