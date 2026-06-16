# -*- coding: utf-8 -*-
"""
test_pages.py — ทดสอบว่าไม่มีหน้ากระดาษเปล่าคั่นเวลาสร้างหลายคน
ต้องมี Microsoft Word (ใช้แปลงเป็น PDF เพื่อนับหน้าจริง)
รัน:  py -3.12 test_pages.py
"""
import os
import merge_engine as me
import word_io as wio
import fitz


def page_report(docx):
    pdf = wio.export_pdf(docx, os.path.splitext(docx)[0] + ".pdf")
    d = fitz.open(pdf)
    n = d.page_count
    blanks = [i + 1 for i, pg in enumerate(d) if not pg.get_text().strip()]
    d.close()
    return n, blanks


def check(label, docx, want_pages):
    n, blanks = page_report(docx)
    ok = (n == want_pages) and (not blanks)
    print(f"[{'PASS' if ok else 'FAIL'}] {label}: {n} หน้า (คาดหวัง {want_pages}), หน้าเปล่า={blanks}")
    return ok


def main():
    os.makedirs("output", exist_ok=True)
    all_ok = True

    # type1 หลายคน (เคสที่เคยมีหน้าเปล่า)
    rows = [{"name": f"คน {i}", "position": f"ตำแหน่ง {i}", "org": f"หน่วยงาน {i}"} for i in range(1, 5)]
    p = me.generate_batch(me.TYPE1, rows, r"output\_test_t1_multi.docx", separate=False)[0]
    all_ok &= check("type1 รวมไฟล์เดียว 4 คน", p, 4)

    # type2 หลายคน
    rows2 = [{"name": f"ชื่อ {i}"} for i in range(1, 4)]
    p = me.generate_batch(me.TYPE2, rows2, r"output\_test_t2_multi.docx", separate=False)[0]
    all_ok &= check("type2 รวมไฟล์เดียว 3 คน", p, 3)

    # type1 ทีละคน
    p = me.generate_one(me.TYPE1, "เดี่ยว", "ตำแหน่ง", "หน่วยงาน", out_path=r"output\_test_t1_one.docx")
    all_ok &= check("type1 ทีละคน", p, 1)

    # type1 หลายคน + หน่วยงานว่างบางคน
    rows3 = [
        {"name": "ก", "position": "ตน ก", "org": "นง ก"},
        {"name": "ข", "position": "ตน ข", "org": ""},
        {"name": "ค", "position": "ตน ค", "org": "นง ค"},
    ]
    p = me.generate_batch(me.TYPE1, rows3, r"output\_test_t1_mixed.docx", separate=False)[0]
    all_ok &= check("type1 หลายคน (หน่วยงานว่างบางคน)", p, 3)

    print("\n=== " + ("ทุกเทสต์ผ่าน ✓" if all_ok else "มีเทสต์ไม่ผ่าน ✗") + " ===")
    return 0 if all_ok else 1


if __name__ == "__main__":
    raise SystemExit(main())
