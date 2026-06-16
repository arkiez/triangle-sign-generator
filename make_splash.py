# -*- coding: utf-8 -*-
"""สร้างรูปหน้าโหลด splash.png สำหรับ PyInstaller (วาดด้วย Pillow)
ใช้ตอน build เท่านั้น ไม่ใช่ runtime dependency ของโปรแกรม

ฝังข้อความไทย "กำลังโหลด…" ลงในรูปเลย (ไม่พึ่ง pyi_splash.update_text
เพราะ Tk-splash ของ PyInstaller อาจไม่มีฟอนต์ไทย) แล้วให้ app.py สั่ง pyi_splash.close()
เมื่อหน้าต่างหลักพร้อม
"""
import os
import re

from PIL import Image, ImageDraw, ImageFont

BASE = os.path.dirname(os.path.abspath(__file__))
FONT_DIR = os.path.join(BASE, "fonts")
KANIT_BOLD = os.path.join(FONT_DIR, "Kanit-Bold.ttf")
KANIT_REG = os.path.join(FONT_DIR, "Kanit-Regular.ttf")

W, H = 480, 300
SS = 2  # วาดใหญ่ x2 แล้วย่อ เพื่อขอบเรียบ (supersampling)


def _app_version(default="1.3.0"):
    """อ่านเลขเวอร์ชันจาก app.py ให้ splash ตรงกับโปรแกรมเสมอ"""
    try:
        with open(os.path.join(BASE, "app.py"), encoding="utf-8") as f:
            m = re.search(r'APP_VERSION\s*=\s*"([^"]+)"', f.read())
        if m:
            return m.group(1)
    except Exception:
        pass
    return default


def _lerp(a, b, t):
    return tuple(int(a[i] + (b[i] - a[i]) * t) for i in range(3))


def _font(path, size):
    try:
        return ImageFont.truetype(path, size)
    except Exception:
        return ImageFont.load_default()


def _fit_font(path, text, max_w, start, min_size):
    """เลือกขนาดฟอนต์ใหญ่สุดที่ข้อความยังกว้างไม่เกิน max_w"""
    size = start
    while size > min_size:
        f = _font(path, size)
        if f.getbbox(text)[2] <= max_w:
            return f
        size -= 1
    return _font(path, min_size)


def _draw_prism(d, ox, oy, s):
    """วาดโลโก้ป้ายตั้งสามเหลี่ยม (ปริซึม 3 มิติ) — logic เดียวกับ make_icon.py
    ในกรอบขนาด s เริ่มที่มุม (ox, oy)  (ไม่มีพื้นหลังการ์ด)"""
    def p(fx, fy):
        return (ox + fx * s, oy + fy * s)

    apex_f = p(0.42, 0.22)   # ยอดหน้า
    base_l = p(0.22, 0.74)   # ฐานซ้าย (หน้า)
    base_r = p(0.62, 0.74)   # ฐานขวา (หน้า)
    apex_b = p(0.60, 0.16)   # ยอดหลัง
    base_br = p(0.80, 0.68)  # ฐานขวาหลัง

    # เงาใต้ป้าย
    d.ellipse([ox + 0.20 * s, oy + 0.74 * s, ox + 0.78 * s, oy + 0.84 * s], fill=(0, 0, 0, 70))
    # หน้าข้าง (เข้มกว่า ให้ดู 3 มิติ)
    d.polygon([apex_f, apex_b, base_br, base_r], fill=(0xDC, 0xE6, 0xF2, 255))
    # สันบน
    d.line([apex_f, apex_b], fill=(0xB8, 0xC8, 0xDC, 255), width=int(s * 0.012))
    # หน้าหน้า (ขาว)
    d.polygon([apex_f, base_l, base_r], fill=(0xFF, 0xFF, 0xFF, 255))
    # เส้นพับกลาง
    mid_base = ((base_l[0] + base_r[0]) / 2, (base_l[1] + base_r[1]) / 2)
    d.line([apex_f, mid_base], fill=(0xCF, 0xDA, 0xE8, 255), width=int(s * 0.010))
    # แถบ "ชื่อ" บนหน้าป้าย (ส้มให้เด่น + เทา)
    d.rounded_rectangle([ox + 0.30 * s, oy + 0.55 * s, ox + 0.545 * s, oy + 0.605 * s],
                        radius=int(s * 0.012), fill=(0xF5, 0x9E, 0x0B, 255))
    d.rounded_rectangle([ox + 0.335 * s, oy + 0.64 * s, ox + 0.515 * s, oy + 0.675 * s],
                        radius=int(s * 0.010), fill=(0x64, 0x74, 0x8B, 255))


def draw():
    w, h = W * SS, H * SS
    img = Image.new("RGBA", (w, h), (0, 0, 0, 0))

    # ---- พื้นหลัง: ไล่สีน้ำเงิน เต็มภาพ (ไม่โปร่งใส กันกรอบดำบน Windows) ----
    top, bot = (0x2B, 0x6C, 0xB0), (0x1B, 0x44, 0x6B)
    grad = Image.new("RGB", (1, h))
    for y in range(h):
        grad.putpixel((0, y), _lerp(top, bot, y / h))
    img.paste(grad.resize((w, h)), (0, 0))
    d = ImageDraw.Draw(img)

    # ---- โลโก้ปริซึม (ซ้าย) ----
    prism = 210 * SS
    _draw_prism(d, ox=18 * SS, oy=58 * SS, s=prism)

    # ---- ข้อความ (ขวา) ----
    tx = 224 * SS
    max_tw = (W - 224 - 18) * SS
    ver = _app_version()

    f_title = _fit_font(KANIT_BOLD, "สร้างป้ายสามเหลี่ยม", max_tw, 34 * SS, 18 * SS)
    d.text((tx, 108 * SS), "สร้างป้ายสามเหลี่ยม", font=f_title, fill=(0xFF, 0xFF, 0xFF, 255))

    f_sub = _font(KANIT_REG, 16 * SS)
    d.text((tx, 150 * SS), f"Mail Merge · v{ver}", font=f_sub, fill=(0xCB, 0xDD, 0xF2, 255))

    # ---- ข้อความ "กำลังโหลด…" (ล่าง, กลางภาพ) — ไม่มีแถบ % ปลอม ----
    f_load = _font(KANIT_REG, 18 * SS)
    txt = "กำลังโหลด…"
    tw = f_load.getbbox(txt)[2]
    d.text(((w - tw) // 2, 255 * SS), txt, font=f_load, fill=(0xE5, 0xEE, 0xFA, 255))

    return img.resize((W, H), Image.LANCZOS)


def main():
    img = draw().convert("RGB")  # บันทึกเป็น RGB (ไม่ต้องมี alpha) เพื่อใช้กับ PyInstaller splash
    out = os.path.join(BASE, "splash.png")
    img.save(out)
    print("saved", out)


if __name__ == "__main__":
    main()
