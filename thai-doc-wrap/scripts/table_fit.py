"""
table_fit.py — ป้องกันตารางล้นขอบกระดาษใน .docx

ปัญหา: ตารางที่กำหนดความกว้างคอลัมน์แบบ fixed หรือ copy-paste จาก Excel
มักล้นออกนอกขอบกระดาษ เห็นเป็นแถบสีเทาหรือถูกตัดออกเมื่อพิมพ์

วิธีแก้ที่ script นี้ทำ:
  1. ตรวจทุกตารางในเอกสาร
  2. คำนวณพื้นที่ข้อความ = ความกว้างหน้า − ขอบซ้าย − ขอบขวา
  3. ถ้าตารางกว้างกว่าพื้นที่ข้อความ → ย่อสัดส่วนทุกคอลัมน์ให้พอดี
  4. ถ้าตารางไม่มีความกว้างกำหนด → ตั้ง AutoFit to Window

การใช้งาน:
  python3 scripts/table_fit.py input.docx -o output.docx
  python3 scripts/table_fit.py input.docx --mode autofit    # ใช้ AutoFit ทุกตาราง
  python3 scripts/table_fit.py input.docx --mode fixed      # ย่อสัดส่วน (ไม่ AutoFit)
  python3 scripts/table_fit.py input.docx --mode pct        # ตั้งเป็น % ของพื้นที่ข้อความ
"""

import argparse
import sys
from copy import deepcopy
from pathlib import Path

try:
    from docx import Document
    from docx.oxml.ns import qn
    from docx.oxml import OxmlElement
    from docx.shared import Twips
    import lxml.etree as etree
except ImportError:
    sys.exit("ต้องติดตั้ง python-docx ก่อน:  pip install python-docx")


def get_text_width_twips(doc) -> int:
    """คำนวณความกว้างพื้นที่ข้อความ (twips) จาก section แรก"""
    sec = doc.sections[0]
    page_w  = sec.page_width.twips  if sec.page_width  else 12240  # A4 ≈ 11906
    left_m  = sec.left_margin.twips if sec.left_margin else 1800
    right_m = sec.right_margin.twips if sec.right_margin else 1800
    return page_w - left_m - right_m


def get_table_width_twips(tbl) -> int | None:
    """อ่านความกว้างรวมของตารางจาก XML (None = ไม่กำหนด)"""
    tblPr  = tbl._tbl.find(qn("w:tblPr"))
    if tblPr is None:
        return None
    tblW = tblPr.find(qn("w:tblW"))
    if tblW is None:
        return None
    w_type = tblW.get(qn("w:type"), "auto")
    w_val  = tblW.get(qn("w:w"), "0")
    if w_type == "dxa":
        return int(w_val)
    if w_type == "pct":
        return None  # % จัดการแยก
    return None


def get_col_widths_twips(tbl) -> list[int]:
    """อ่านความกว้างแต่ละคอลัมน์ (twips) จาก tblGrid"""
    grid = tbl._tbl.find(qn("w:tblGrid"))
    if grid is None:
        return []
    cols = []
    for gc in grid.findall(qn("w:gridCol")):
        w = gc.get(qn("w:w"), "0")
        cols.append(int(w))
    return cols


def set_col_widths(tbl, new_widths: list[int]):
    """ตั้งความกว้างคอลัมน์ใหม่ใน tblGrid และแต่ละ cell"""
    grid = tbl._tbl.find(qn("w:tblGrid"))
    if grid is not None:
        grid_cols = grid.findall(qn("w:gridCol"))
        for gc, w in zip(grid_cols, new_widths):
            gc.set(qn("w:w"), str(w))

    # อัปเดต cell width ในแต่ละแถว
    for row in tbl.rows:
        cells = row.cells
        for i, cell in enumerate(cells):
            if i >= len(new_widths):
                break
            tc  = cell._tc
            tcPr = tc.find(qn("w:tcPr"))
            if tcPr is None:
                tcPr = OxmlElement("w:tcPr")
                tc.insert(0, tcPr)
            tcW = tcPr.find(qn("w:tcW"))
            if tcW is None:
                tcW = OxmlElement("w:tcW")
                tcPr.append(tcW)
            tcW.set(qn("w:type"), "dxa")
            tcW.set(qn("w:w"), str(new_widths[i]))


def set_table_autofit(tbl):
    """ตั้ง AutoFit to Window"""
    tblPr = tbl._tbl.find(qn("w:tblPr"))
    if tblPr is None:
        tblPr = OxmlElement("w:tblPr")
        tbl._tbl.insert(0, tblPr)

    # ตั้ง tblW เป็น 100% (pct = 5000 = 100%)
    tblW = tblPr.find(qn("w:tblW"))
    if tblW is None:
        tblW = OxmlElement("w:tblW")
        tblPr.append(tblW)
    tblW.set(qn("w:type"), "pct")
    tblW.set(qn("w:w"), "5000")

    # ลบ tblLayout fixed ถ้ามี
    tblLayout = tblPr.find(qn("w:tblLayout"))
    if tblLayout is not None:
        tblPr.remove(tblLayout)


def process_doc(doc_path: Path, out_path: Path, mode: str) -> dict:
    doc = Document(str(doc_path))
    text_w = get_text_width_twips(doc)
    stats = {"checked": 0, "scaled": 0, "autofit": 0, "ok": 0}

    for tbl in doc.tables:
        stats["checked"] += 1

        if mode == "autofit":
            set_table_autofit(tbl)
            stats["autofit"] += 1
            continue

        col_widths = get_col_widths_twips(tbl)
        tbl_total  = sum(col_widths) if col_widths else None

        if mode == "pct":
            set_table_autofit(tbl)
            stats["autofit"] += 1
            continue

        # mode == "fixed" หรือ "auto"
        if not col_widths or tbl_total is None or tbl_total == 0:
            # ไม่มีข้อมูลความกว้าง → AutoFit
            set_table_autofit(tbl)
            stats["autofit"] += 1
        elif tbl_total > text_w:
            # ล้นขอบ → ย่อสัดส่วน
            ratio = text_w / tbl_total
            new_widths = [max(300, int(w * ratio)) for w in col_widths]
            set_col_widths(tbl, new_widths)

            # อัปเดต tblW ให้ตรง
            tblPr = tbl._tbl.find(qn("w:tblPr"))
            if tblPr is not None:
                tblW = tblPr.find(qn("w:tblW"))
                if tblW is not None:
                    tblW.set(qn("w:type"), "dxa")
                    tblW.set(qn("w:w"), str(sum(new_widths)))
            stats["scaled"] += 1
        else:
            stats["ok"] += 1

    doc.save(str(out_path))
    return stats, text_w


def main():
    ap = argparse.ArgumentParser(description="ป้องกันตารางล้นขอบใน .docx")
    ap.add_argument("input", help="ไฟล์ .docx ต้นทาง")
    ap.add_argument("-o", "--output", help="ไฟล์ผลลัพธ์ (ไม่ระบุ = เพิ่ม .tablefit.docx)")
    ap.add_argument("--mode", choices=["auto", "autofit", "fixed", "pct"], default="auto",
                    help="auto=ย่อเมื่อล้น (ค่าเริ่มต้น), autofit=AutoFit ทุกตาราง, "
                         "fixed=ย่อสัดส่วนเสมอ, pct=100%% เสมอ")
    args = ap.parse_args()

    src = Path(args.input)
    if not src.exists():
        sys.exit(f"ไม่พบไฟล์: {src}")

    dst = Path(args.output) if args.output else src.with_suffix(".tablefit.docx")
    stats, text_w = process_doc(src, dst, args.mode)

    import io
    out = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8", errors="replace")
    out.write(f"OK: checked {stats['checked']} table(s) | scaled: {stats['scaled']} | "
              f"autofit: {stats['autofit']} | ok: {stats['ok']}\n")
    out.write(f"   text area: {text_w/20:.1f} pt ({text_w/1440*2.54:.2f} cm)\n")
    out.write(f"   saved -> {dst}\n")
    out.flush()


if __name__ == "__main__":
    main()
