#!/usr/bin/env python3
"""
thai_wrap.py — แทรกจุดตัดบรรทัดที่ขอบเขตคำภาษาไทย เพื่อให้เอกสารตัดคำสวยงาม
ไม่ตัดกลางคำ

หลักการ: ภาษาไทยไม่เว้นวรรคระหว่างคำ โปรแกรมจัดหน้า (เบราว์เซอร์/Word/PDF)
จึงไม่รู้ขอบเขตคำและมักตัดบรรทัดกลางคำ วิธีแก้คือใช้พจนานุกรมตัดคำ (PyThaiNLP)
หาขอบเขตคำ แล้วแทรกอักขระ Zero-Width Space (U+200B) ซึ่งมองไม่เห็นแต่บอก
โปรแกรมว่า "ตัดบรรทัดตรงนี้ได้" — ผลคือบรรทัดจะตัดที่ท้ายคำเสมอ

รองรับ: .txt, .html/.htm, .docx, และ stdin/stdout
ขึ้นกับ: pythainlp (จำเป็น), beautifulsoup4 (สำหรับ HTML), python-docx (สำหรับ DOCX)
"""
import argparse
import sys
import os

ZWSP = "\u200b"   # Zero-Width Space — จุดตัดบรรทัดที่มองไม่เห็น
WJ = "\u2060"     # Word Joiner — กันไม่ให้ตัดบรรทัด (ใช้ในกรณีพิเศษ)

# อักขระที่ไม่ควรขึ้นต้นบรรทัด (วรรณยุกต์ ไม้ยมก ฯลฯ) — ไม่แทรกจุดตัด "ก่อน" อักขระเหล่านี้
THAI_NO_LINE_START = set("ๆฯ")


def _strip_existing(text: str) -> str:
    """ลบ ZWSP/WJ เดิมออกก่อน เพื่อให้ทำซ้ำได้โดยไม่สะสมอักขระ (idempotent)"""
    return text.replace(ZWSP, "").replace(WJ, "")


def segment_text(text: str, engine: str = "newmm", marker: str = ZWSP) -> str:
    """ตัดคำข้อความไทยแล้วคืนข้อความเดิมที่แทรก marker ที่ขอบเขตคำ

    แทรกเฉพาะระหว่างโทเคนที่ไม่ใช่ช่องว่าง และไม่แทรกก่อนอักขระที่ห้ามขึ้นต้นบรรทัด
    """
    from pythainlp.tokenize import word_tokenize

    if not text or not text.strip():
        return text
    text = _strip_existing(text)
    tokens = word_tokenize(text, engine=engine, keep_whitespace=True)

    out = []
    for i, tok in enumerate(tokens):
        if i > 0 and tok:
            prev = tokens[i - 1]
            prev_ok = prev and not prev[-1].isspace()
            cur_ok = not tok[0].isspace()
            not_orphan_mark = tok[0] not in THAI_NO_LINE_START
            if prev_ok and cur_ok and not_orphan_mark:
                out.append(marker)
        out.append(tok)
    return "".join(out)


# ---------- โหมดข้อความล้วน ----------
def process_text_file(text: str, engine: str, marker: str) -> str:
    lines = text.split("\n")
    return "\n".join(segment_text(ln, engine, marker) for ln in lines)


# ---------- โหมด HTML ----------
SKIP_TAGS = {"script", "style", "pre", "code", "textarea", "kbd", "samp"}


def process_html(html: str, engine: str, marker_mode: str) -> str:
    """แทรกจุดตัดเฉพาะใน text node ของ HTML โดยไม่แตะ markup
    marker_mode: 'zwsp' แทรก U+200B | 'wbr' แทรกแท็ก <wbr>
    """
    from bs4 import BeautifulSoup, NavigableString

    soup = BeautifulSoup(html, "html.parser")

    for node in list(soup.find_all(string=True)):
        if not isinstance(node, NavigableString):
            continue
        if node.parent and node.parent.name in SKIP_TAGS:
            continue
        original = str(node)
        if not original.strip():
            continue
        if marker_mode == "wbr":
            # แตกข้อความออกเป็นชิ้น คั่นด้วยแท็ก <wbr>
            seg = segment_text(original, engine, marker="\x00")
            parts = seg.split("\x00")
            if len(parts) <= 1:
                continue
            new_nodes = []
            for j, part in enumerate(parts):
                if j > 0:
                    new_nodes.append(soup.new_tag("wbr"))
                new_nodes.append(NavigableString(part))
            node.replace_with(*new_nodes)
        else:
            node.replace_with(NavigableString(segment_text(original, engine, ZWSP)))

    return str(soup)


# ---------- โหมด DOCX ----------
def _iter_paragraphs(doc):
    """วนทุกย่อหน้าในเอกสาร รวมในตาราง หัวกระดาษ และท้ายกระดาษ"""
    from docx.table import Table
    from docx.text.paragraph import Paragraph

    def walk(container):
        for p in getattr(container, "paragraphs", []):
            yield p
        for t in getattr(container, "tables", []):
            for row in t.rows:
                for cell in row.cells:
                    yield from walk(cell)

    yield from walk(doc)
    for section in doc.sections:
        for hf in (section.header, section.footer,
                   section.first_page_header, section.first_page_footer,
                   section.even_page_header, section.even_page_footer):
            yield from walk(hf)


def _set_run_thai_lang(run):
    """ตั้งค่าภาษาของ run เป็นไทย(complex script) เพื่อให้ตัวตัดคำของ Word ทำงานร่วมด้วย"""
    from docx.oxml.ns import qn
    rpr = run._element.get_or_add_rPr()
    lang = rpr.find(qn("w:lang"))
    if lang is None:
        lang = rpr.makeelement(qn("w:lang"), {})
        rpr.append(lang)
    lang.set(qn("w:bidi"), "th-TH")


def _para_is_justified(para) -> bool:
    """คืนค่า True ถ้าย่อหน้ากำหนด text-align: justify (w:jc = both/distribute)
    Word กระจาย space ตาม ZWSP ทุกจุดในโหมด justify ทำให้ช่องว่างถ่างมาก
    จึงต้องข้ามการแทรก ZWSP และใช้ th-TH language ให้ Word ตัดคำเองแทน
    """
    from docx.oxml.ns import qn
    pPr = para._p.find(qn("w:pPr"))
    if pPr is None:
        return False
    jc = pPr.find(qn("w:jc"))
    if jc is None:
        return False
    val = jc.get(qn("w:val"), "")
    return val in ("both", "distribute", "highKashida", "lowKashida", "mediumKashida")


def process_docx(in_path: str, out_path: str, engine: str, set_lang: bool = True) -> int:
    """แทรก ZWSP ใน run ที่ไม่ใช่ justified paragraph
    สำหรับ justified paragraph: ข้าม ZWSP แต่ยังตั้ง th-TH language เพื่อให้ Word
    ใช้ตัวตัดคำในตัวเอง (ซึ่งทำงานได้ดีกับ justify โดยไม่เกิดช่องว่างถ่าง)
    """
    from docx import Document

    doc = Document(in_path)
    n_runs = 0
    n_skipped_justify = 0
    for para in _iter_paragraphs(doc):
        justified = _para_is_justified(para)
        for run in para.runs:
            if not run.text or not run.text.strip():
                continue
            if set_lang:
                _set_run_thai_lang(run)
            if justified:
                # justified → ไม่แทรก ZWSP ให้ Word ตัดคำเองผ่าน th-TH lang
                n_skipped_justify += 1
            else:
                run.text = segment_text(run.text, engine, ZWSP)
                n_runs += 1
    doc.save(out_path)
    return n_runs, n_skipped_justify


def main():
    ap = argparse.ArgumentParser(
        description="แทรกจุดตัดบรรทัดที่ขอบเขตคำไทย เพื่อไม่ให้ตัดกลางคำ")
    ap.add_argument("input", nargs="?", help="ไฟล์เข้า (.txt/.html/.docx) หรือเว้นไว้เพื่ออ่านจาก stdin")
    ap.add_argument("-o", "--output", help="ไฟล์ออก (เว้นไว้ = stdout สำหรับ text/html)")
    ap.add_argument("--engine", default="newmm",
                    help="เอนจินตัดคำ PyThaiNLP: newmm (เร็ว, ค่าเริ่มต้น), longest, nlpo3")
    ap.add_argument("--format", choices=["auto", "text", "html", "docx"], default="auto",
                    help="บังคับรูปแบบไฟล์ (ค่าเริ่มต้น auto ตามนามสกุล)")
    ap.add_argument("--marker", choices=["zwsp", "wbr"], default="zwsp",
                    help="HTML: zwsp=U+200B (ค่าเริ่มต้น, ใช้ได้กับ PDF) | wbr=แท็ก <wbr>")
    ap.add_argument("--no-lang", action="store_true",
                    help="DOCX: ไม่ต้องตั้งค่าภาษา complex-script เป็นไทย")
    args = ap.parse_args()

    fmt = args.format
    if fmt == "auto":
        if args.input:
            ext = os.path.splitext(args.input)[1].lower()
            fmt = {"docx": "docx", "html": "html", "htm": "html"}.get(ext.lstrip("."), "text")
        else:
            fmt = "text"

    if fmt == "docx":
        if not args.input:
            sys.exit("DOCX ต้องระบุไฟล์เข้า")
        out = args.output or args.input.rsplit(".", 1)[0] + ".wrapped.docx"
        n, n_skip = process_docx(args.input, out, args.engine, set_lang=not args.no_lang)
        print(f"✓ ตัดคำแล้ว {n} run (ข้าม {n_skip} run ใน justified para) → {out}", file=sys.stderr)
        return

    data = open(args.input, encoding="utf-8").read() if args.input else sys.stdin.read()
    if fmt == "html":
        result = process_html(data, args.engine, args.marker)
    else:
        result = process_text_file(data, args.engine, ZWSP)

    if args.output:
        with open(args.output, "w", encoding="utf-8") as f:
            f.write(result)
        print(f"✓ เขียน → {args.output}", file=sys.stderr)
    else:
        sys.stdout.write(result)


if __name__ == "__main__":
    main()
