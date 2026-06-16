# -*- coding: utf-8 -*-
"""
merge_engine.py
เอนจินเติมข้อมูลลง Word template ป้ายสามเหลี่ยม (MERGEFIELD) -> ไฟล์ .docx
- ใช้ docx-mailmerge2 เป็นตัวหลัก (รักษาเลย์เอาต์/ฟอนต์/การหมุน 180° ของ template ไว้ครบ)
- ถ้าไม่มี docx-mailmerge2 จะ fallback ไปใช้ zipfile แทนข้อความ placeholder แบบง่าย
ทำงานได้โดย "ไม่ต้องมี Microsoft Word" (Word ใช้เฉพาะตอนพรีวิว/PDF/พิมพ์ ใน word_io.py)
"""
import os
import re
import shutil
import sys
import tempfile
import zipfile

# ---- ค่าคงที่ของชนิดป้าย ----
TYPE1 = "type1_name_pos"   # ชื่อ + ตำแหน่ง/หน่วยงาน
TYPE2 = "type2_name"       # ชื่ออย่างเดียว

# รองรับทั้งรันจากซอร์ส และรันจาก .exe (PyInstaller)
if getattr(sys, "frozen", False):
    BASE = sys._MEIPASS                          # ทรัพยากร (templates) ถูกแตกไว้ที่นี่
    APP_DIR = os.path.dirname(sys.executable)    # โฟลเดอร์ที่วาง .exe (เขียนไฟล์ได้)
else:
    BASE = os.path.dirname(os.path.abspath(__file__))
    APP_DIR = BASE
TEMPLATES_DIR = os.path.join(BASE, "templates")
OUTPUT_DIR = os.path.join(APP_DIR, "output")

TEMPLATE_FILES = {
    TYPE1: os.path.join(TEMPLATES_DIR, "type1_name_pos.docx"),
    TYPE2: os.path.join(TEMPLATES_DIR, "type2_name.docx"),
}

TYPE_LABELS = {
    TYPE1: "แบบ 1: ชื่อ + ตำแหน่ง/หน่วยงาน",
    TYPE2: "แบบ 2: ชื่ออย่างเดียว",
}


# ---------------------------------------------------------------------------
# ตัวช่วย
# ---------------------------------------------------------------------------
def build_pos(position: str, org: str) -> str:
    """รวม 'ตำแหน่ง' + 'หน่วยงาน' เป็นค่า pos เดียว คั่นด้วยขึ้นบรรทัดใหม่
    ถ้าหน่วยงานว่าง จะไม่ใส่บรรทัดเปล่า"""
    position = (position or "").strip()
    org = (org or "").strip()
    if position and org:
        return position + "\n" + org
    return position or org


def sanitize_filename(name: str) -> str:
    """ตัดอักขระต้องห้ามของ Windows ออกจากชื่อไฟล์ (รองรับภาษาไทย)"""
    name = (name or "").strip()
    name = re.sub(r'[\\/:*?"<>|\r\n\t]+', " ", name)
    name = re.sub(r"\s+", " ", name).strip()
    # ตัดจุด/ช่องว่างท้ายชื่อ (Windows ไม่ชอบ)
    name = name.rstrip(". ")
    return name or "ป้าย"


def unique_path(path: str) -> str:
    """ถ้าไฟล์ซ้ำ เติม (2), (3), ... กันทับของเดิม"""
    if not os.path.exists(path):
        return path
    base, ext = os.path.splitext(path)
    i = 2
    while os.path.exists(f"{base} ({i}){ext}"):
        i += 1
    return f"{base} ({i}){ext}"


def template_path(template_type: str) -> str:
    p = TEMPLATE_FILES.get(template_type)
    if not p or not os.path.exists(p):
        raise FileNotFoundError(
            f"ไม่พบไฟล์ template สำหรับ {template_type}: {p}\n"
            f"กรุณาตรวจสอบโฟลเดอร์ templates\\"
        )
    return p


def _row_to_data(template_type: str, row: dict) -> dict:
    """แปลงข้อมูล 1 แถว -> dict ของ merge field"""
    data = {"name": (row.get("name") or "").strip()}
    if template_type == TYPE1:
        data["pos"] = build_pos(row.get("position", ""), row.get("org", ""))
    return data


# ---------------------------------------------------------------------------
# API หลัก
# ---------------------------------------------------------------------------
def generate_one(template_type, name, position="", org="", out_path=None) -> str:
    """สร้างป้าย 1 ใบ (1 คน = 1 หน้า, เติมครบทั้ง 4 จุดด้วยค่าเดียวกัน)
    คืนค่า: path ของไฟล์ .docx ที่สร้าง"""
    tpl = template_path(template_type)
    data = _row_to_data(template_type, {"name": name, "position": position, "org": org})

    if out_path is None:
        os.makedirs(OUTPUT_DIR, exist_ok=True)
        out_path = unique_path(os.path.join(OUTPUT_DIR, sanitize_filename(name) + ".docx"))

    _merge_to_file(tpl, [data], out_path)
    return out_path


def generate_batch(template_type, rows, out_path, separate=False):
    """สร้างป้ายหลายใบจากรายการ rows (list ของ dict {name, position, org})
    - separate=False : รวมเป็นไฟล์เดียว (1 คน = 1 หน้า) -> out_path เป็นไฟล์ .docx
    - separate=True  : แยกไฟล์ต่อคน -> out_path เป็นโฟลเดอร์
    คืนค่า: list ของ path ที่สร้าง"""
    tpl = template_path(template_type)
    rows = [r for r in rows if (r.get("name") or "").strip()]  # ข้ามแถวที่ไม่มีชื่อ
    if not rows:
        raise ValueError("ไม่มีข้อมูล (ทุกแถวไม่มีชื่อ)")

    datas = [_row_to_data(template_type, r) for r in rows]

    if separate:
        os.makedirs(out_path, exist_ok=True)
        paths = []
        for r, d in zip(rows, datas):
            p = unique_path(os.path.join(out_path, sanitize_filename(r.get("name")) + ".docx"))
            _merge_to_file(tpl, [d], p)
            paths.append(p)
        return paths

    os.makedirs(os.path.dirname(os.path.abspath(out_path)), exist_ok=True)
    _merge_to_file(tpl, datas, out_path)
    return [out_path]


# ---------------------------------------------------------------------------
# ตัวเติมจริง: docx-mailmerge2 (หลัก) + fallback
# ---------------------------------------------------------------------------
def _merge_to_file(tpl, datas, out_path):
    os.makedirs(os.path.dirname(os.path.abspath(out_path)), exist_ok=True)
    try:
        from mailmerge import MailMerge
    except ImportError:
        return _merge_fallback(tpl, datas, out_path)

    # กัน "หน้ากระดาษเปล่าคั่นระหว่างคน" ตอน merge หลายคน:
    # template บางไฟล์มี section break (sectPr) ฝังในย่อหน้า ทำให้ขึ้นหน้าใหม่เอง
    # เมื่อรวมกับตัวคั่น page_break ของ mail merge จะกลายเป็นหน้าเปล่า 1 หน้า
    tpl = _normalized_template(tpl)

    with MailMerge(tpl) as mm:
        if len(datas) == 1:
            # เติมค่าเดียว -> ทุก field ชื่อเดียวกัน (name x4, pos x4) ได้ค่าเดียวกัน
            mm.merge(**datas[0])
        else:
            # หลายคน: 1 คน = 1 หน้า (ขึ้นหน้าใหม่)
            mm.merge_templates(datas, separator="page_break")
        mm.write(out_path)

    _autoshrink(out_path)  # ย่อฟอนต์อัตโนมัติถ้าข้อความยาวเกินกล่อง
    return out_path


# ---------------------------------------------------------------------------
# ย่อฟอนต์อัตโนมัติ: วัดความกว้างข้อความจริง ถ้ายาวเกินกล่องข้อความ -> ลดขนาดให้พอดี
# ---------------------------------------------------------------------------
_W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
_A = "{http://schemas.openxmlformats.org/drawingml/2006/main}"
_WPS = "{http://schemas.microsoft.com/office/word/2010/wordprocessingShape}"
_EMU_PER_PT = 12700.0
_SHRINK_FLOOR = 0.40        # ไม่ย่อเล็กกว่า 40% ของขนาดเดิม
_MIN_HALFPT = 32            # และไม่เล็กกว่า 16pt เด็ดขาด

# สระ/วรรณยุกต์ที่ไม่กินความกว้าง (ใช้ตอน fallback ไม่มีฟอนต์วัด)
_THAI_COMBINING = set("่้๊๋ัิีึืุู็์ํๅ".replace("ๅ", "")) | set("ัิีึืฺุู็่้๊๋์ํ๎")
_font_cache = {}


def _find_sarabun_ttf(bold):
    """หาไฟล์ TTF ของ TH Sarabun (ตัวหนา/ปกติ) จากโฟลเดอร์ฟอนต์ของ Windows"""
    import glob
    dirs = [
        os.path.join(os.environ.get("WINDIR", r"C:\Windows"), "Fonts"),
        os.path.join(os.environ.get("LOCALAPPDATA", ""), "Microsoft", "Windows", "Fonts"),
    ]
    want_bold = ["THSarabun Bold.ttf", "THSarabunNew Bold.ttf", "THSarabunPSK Bold.ttf"]
    want_reg = ["THSarabun.ttf", "THSarabunNew.ttf", "THSarabunPSK.ttf"]
    wants = want_bold if bold else want_reg
    for d in dirs:
        if not d or not os.path.isdir(d):
            continue
        for name in wants:
            p = os.path.join(d, name)
            if os.path.exists(p):
                return p
        # เผื่อชื่อไฟล์ต่างไป
        hits = glob.glob(os.path.join(d, "*[Ss]arabun*.ttf"))
        if hits:
            bolds = [h for h in hits if "bold" in os.path.basename(h).lower()]
            return (bolds or hits)[0] if bold else (([h for h in hits if "bold" not in os.path.basename(h).lower()] or hits)[0])
    return None


def _get_font(bold):
    key = bool(bold)
    if key not in _font_cache:
        f = None
        try:
            import fitz
            path = _find_sarabun_ttf(bold)
            if path:
                f = fitz.Font(fontfile=path)
        except Exception:
            f = None
        _font_cache[key] = f
    return _font_cache[key]


def _text_width_pt(text, pt, bold):
    """ความกว้างข้อความเป็น point ที่ขนาดฟอนต์ pt"""
    if not text:
        return 0.0
    f = _get_font(bold)
    if f is not None:
        try:
            return f.text_length(text, pt)
        except Exception:
            pass
    # fallback: นับเฉพาะตัวอักษรที่กินความกว้าง (~0.5*pt ต่อตัว)
    base = sum(1 for ch in text if ch not in _THAI_COMBINING)
    return base * 0.5 * pt


def _run_lines(run):
    """ดึงข้อความใน run แยกเป็นบรรทัด (ตาม <w:br/> / <w:cr/>)"""
    from lxml import etree

    lines = [""]
    for child in run:
        tag = etree.QName(child).localname
        if tag == "t":
            lines[-1] += child.text or ""
        elif tag in ("br", "cr"):
            lines.append("")
    return [ln for ln in lines if ln]


def _autoshrink(path):
    """เปิดไฟล์ผลลัพธ์ แล้วย่อขนาดฟอนต์ของ run ที่ข้อความยาวเกินความกว้างกล่องข้อความ"""
    try:
        from lxml import etree
    except ImportError:
        return  # ไม่มี lxml ก็ข้าม (โหมด fallback)

    try:
        with zipfile.ZipFile(path) as z:
            data = z.read("word/document.xml")
        root = etree.fromstring(data)
    except Exception:
        return

    changed = False
    for wsp in root.iter(_WPS + "wsp"):
        ext = wsp.find(".//" + _A + "ext")
        body = wsp.find(".//" + _WPS + "bodyPr")
        txbx = wsp.find(".//" + _W + "txbxContent")
        if ext is None or txbx is None:
            continue
        try:
            cx = int(ext.get("cx"))
        except (TypeError, ValueError):
            continue
        lins = int(body.get("lIns", "91440")) if body is not None else 91440
        rins = int(body.get("rIns", "91440")) if body is not None else 91440
        avail_pt = (cx - lins - rins) / _EMU_PER_PT
        if avail_pt <= 0:
            continue

        for run in txbx.iter(_W + "r"):
            rpr = run.find(_W + "rPr")
            if rpr is None:
                continue
            sz_el = rpr.find(_W + "sz")
            szcs_el = rpr.find(_W + "szCs")
            base_el = szcs_el if szcs_el is not None else sz_el  # Thai ใช้ szCs เป็นหลัก
            if base_el is None:
                continue
            try:
                sz = int(base_el.get(_W + "val"))   # attribute เป็น w:val (มี namespace)
            except (TypeError, ValueError):
                continue
            pt = sz / 2.0
            bold = rpr.find(_W + "b") is not None
            lines = _run_lines(run)
            if not lines:
                continue

            # letter-spacing (w:spacing, หน่วย twips = 1/20 pt ต่อ 1 ตัวอักษร) ที่ fitz วัดไม่เห็น
            spc_el = rpr.find(_W + "spacing")
            try:
                spc_pt = int(spc_el.get(_W + "val")) / 20.0 if spc_el is not None else 0.0
            except (TypeError, ValueError):
                spc_pt = 0.0

            # หาขนาดใหญ่สุดที่ยังพอดีกว้างกล่อง (เผื่อ margin 3%)
            target = avail_pt * 0.97
            limit_pt = pt
            for ln in lines:
                w1 = _text_width_pt(ln, 1.0, bold)   # ความกว้างที่ 1pt (สเกลเชิงเส้น)
                if w1 <= 0:
                    continue
                allowed = (target - spc_pt * len(ln)) / w1   # แก้สมการ w1*pt + spc*n <= target
                if allowed < limit_pt:
                    limit_pt = allowed
            if limit_pt >= pt:
                continue  # พอดีอยู่แล้ว ไม่ต้องย่อ

            floor_pt = max(pt * _SHRINK_FLOOR, _MIN_HALFPT / 2.0)
            new_sz = int(max(limit_pt, floor_pt) * 2)
            if new_sz < sz:
                if sz_el is not None:
                    sz_el.set(_W + "val", str(new_sz))
                if szcs_el is not None:
                    szcs_el.set(_W + "val", str(new_sz))
                changed = True

    if not changed:
        return

    new_doc = etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True)
    tmp = path + ".shrink"
    with zipfile.ZipFile(path) as zin, zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as zout:
        for item in zin.infolist():
            data = new_doc if item.filename == "word/document.xml" else zin.read(item.filename)
            zout.writestr(item, data)
    os.replace(tmp, path)


# แคชผลการ normalize ไว้ใช้ซ้ำ (key = abspath ของ template, value = (mtime, path))
_norm_cache = {}

# namespace ของ WordprocessingML
_W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"


def _normalized_template(tpl):
    """คืน path ของ template ที่เอา section break ภายในย่อหน้า (sectPr ใน pPr) ออกแล้ว
    เพื่อกันหน้ากระดาษเปล่าคั่นเวลา merge หลายคน โดยยังคง sectPr ท้าย body
    (ที่กำหนดขนาด/แนวกระดาษ) ไว้ครบ  เก็บผลใน temp และใช้ซ้ำถ้า template ไม่เปลี่ยน"""
    from lxml import etree

    src = os.path.abspath(tpl)
    mtime = os.path.getmtime(src)
    cached = _norm_cache.get(src)
    if cached and cached[0] == mtime and os.path.exists(cached[1]):
        return cached[1]

    with zipfile.ZipFile(src) as zin:
        doc = zin.read("word/document.xml")

    root = etree.fromstring(doc)
    changed = False
    for pPr in root.iter(_W + "pPr"):
        sect = pPr.find(_W + "sectPr")
        if sect is not None:
            pPr.remove(sect)  # ลบเฉพาะ section break ที่ฝังในย่อหน้า ไม่แตะตัวท้าย body
            changed = True

    if not changed:
        # template นี้ไม่มีปัญหา ใช้ไฟล์เดิมได้เลย
        _norm_cache[src] = (mtime, src)
        return src

    new_doc = etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True)
    dst = os.path.join(tempfile.gettempdir(), "_tent_norm_" + os.path.basename(src))
    tmp = dst + ".tmp"
    with zipfile.ZipFile(src) as zin, zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as zout:
        for item in zin.infolist():
            data = new_doc if item.filename == "word/document.xml" else zin.read(item.filename)
            zout.writestr(item, data)
    os.replace(tmp, dst)
    _norm_cache[src] = (mtime, dst)
    return dst


# ---- Fallback แบบ stdlib ล้วน (เผื่อไม่มี docx-mailmerge2) ----
# จัดการเฉพาะกรณี 1 คน/ไฟล์ โดยแทนข้อความ cached «name»/«pos» ในเลย์เอาต์เดิม
_CR = "</w:t><w:cr/><w:t xml:space=\"preserve\">"


def _xml_escape(text: str) -> str:
    return (
        text.replace("&", "&amp;")
        .replace("<", "&lt;")
        .replace(">", "&gt;")
        .replace('"', "&quot;")
    )


def _to_xml_value(value: str) -> str:
    """escape + แปลง \n เป็น <w:cr/> (ขึ้นบรรทัดในย่อหน้าเดียว)"""
    value = (value or "")
    parts = value.replace("\r", "").split("\n")
    return _CR.join(_xml_escape(p) for p in parts)


def _merge_fallback(tpl, datas, out_path):
    """รองรับ 1 คน/ไฟล์เท่านั้น (กรณีไม่มี docx-mailmerge2)."""
    if len(datas) != 1:
        raise RuntimeError(
            "ไม่ได้ติดตั้ง docx-mailmerge2 จึงสร้างแบบหลายคนในไฟล์เดียวไม่ได้\n"
            "ติดตั้งด้วย: pip install docx-mailmerge2"
        )
    data = datas[0]
    shutil.copyfile(tpl, out_path)

    tmp = out_path + ".tmp"
    with zipfile.ZipFile(out_path, "r") as zin, zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as zout:
        for item in zin.infolist():
            content = zin.read(item.filename)
            if item.filename == "word/document.xml":
                text = content.decode("utf-8")
                for key, val in data.items():
                    text = text.replace(f"«{key}»", _to_xml_value(val))
                # ลบลิงก์ mail merge ใน settings ทำในไฟล์ settings.xml แยก (ด้านล่าง)
                content = text.encode("utf-8")
            elif item.filename == "word/settings.xml":
                text = content.decode("utf-8")
                text = re.sub(r"<w:mailMerge>.*?</w:mailMerge>", "", text, flags=re.DOTALL)
                content = text.encode("utf-8")
            zout.writestr(item, content)
    os.replace(tmp, out_path)
    return out_path


if __name__ == "__main__":
    # ทดสอบเร็ว ๆ
    p1 = generate_one(TYPE1, "นายสมชาย ใจดี", "ผู้อำนวยการ", "สำนักงานเขตพื้นที่การศึกษา")
    print("Type1 ->", p1)
    p2 = generate_one(TYPE2, "นางสาวสมหญิง รักเรียน")
    print("Type2 ->", p2)
