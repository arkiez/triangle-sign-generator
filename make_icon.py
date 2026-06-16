# -*- coding: utf-8 -*-
"""สร้างไอคอน app.ico ธีมป้ายตั้งสามเหลี่ยม (วาดด้วย Pillow)"""
from PIL import Image, ImageDraw

S = 1024  # วาดใหญ่แล้วย่อ เพื่อให้ขอบเรียบ


def lerp(a, b, t):
    return tuple(int(a[i] + (b[i] - a[i]) * t) for i in range(3))


def draw():
    img = Image.new("RGBA", (S, S), (0, 0, 0, 0))

    # ---- พื้นหลัง: ไล่สีน้ำเงิน มุมโค้ง ----
    top, bot = (0x2B, 0x6C, 0xB0), (0x1B, 0x44, 0x6B)
    grad = Image.new("RGB", (1, S))
    for y in range(S):
        grad.putpixel((0, y), lerp(top, bot, y / S))
    grad = grad.resize((S, S))
    mask = Image.new("L", (S, S), 0)
    ImageDraw.Draw(mask).rounded_rectangle([0, 0, S - 1, S - 1], radius=int(S * 0.20), fill=255)
    bg = Image.new("RGBA", (S, S), (0, 0, 0, 0))
    bg.paste(grad, (0, 0), mask)
    img = Image.alpha_composite(img, bg)
    d = ImageDraw.Draw(img)

    # ---- เงาใต้ป้าย ----
    d.ellipse([S * 0.24, S * 0.74, S * 0.76, S * 0.82], fill=(0, 0, 0, 60))

    # ---- ป้ายตั้งสามเหลี่ยม (ปริซึม 3 มิติ) ----
    apex_f = (S * 0.42, S * 0.22)   # ยอดหน้า
    base_l = (S * 0.22, S * 0.74)   # ฐานซ้าย (หน้า)
    base_r = (S * 0.62, S * 0.74)   # ฐานขวา (หน้า)
    apex_b = (S * 0.60, S * 0.16)   # ยอดหลัง (ลึกเข้าไป)
    base_br = (S * 0.80, S * 0.68)  # ฐานขวาหลัง

    # หน้าข้าง (เข้มกว่า ให้ดู 3 มิติ)
    d.polygon([apex_f, apex_b, base_br, base_r], fill=(0xDC, 0xE6, 0xF2, 255))
    # สันบน
    d.line([apex_f, apex_b], fill=(0xB8, 0xC8, 0xDC, 255), width=int(S * 0.012))
    # หน้าหน้า (ขาว)
    d.polygon([apex_f, base_l, base_r], fill=(0xFF, 0xFF, 0xFF, 255))

    # เส้นพับกลาง (จากยอดลงฐาน)
    mid_base = ((base_l[0] + base_r[0]) / 2, (base_l[1] + base_r[1]) / 2)
    d.line([apex_f, mid_base], fill=(0xCF, 0xDA, 0xE8, 255), width=int(S * 0.010))

    # ---- แถบ "ชื่อ" บนหน้าป้าย (สีส้มให้เด่น) ----
    d.rounded_rectangle([S * 0.30, S * 0.55, S * 0.545, S * 0.605],
                        radius=int(S * 0.012), fill=(0xF5, 0x9E, 0x0B, 255))
    d.rounded_rectangle([S * 0.335, S * 0.64, S * 0.515, S * 0.675],
                        radius=int(S * 0.010), fill=(0x64, 0x74, 0x8B, 255))

    return img.resize((256, 256), Image.LANCZOS)


def main():
    icon = draw()
    icon.save("app.png")  # ไว้พรีวิว
    icon.save("app.ico", sizes=[(16, 16), (24, 24), (32, 32), (48, 48),
                                (64, 64), (128, 128), (256, 256)])
    print("saved app.ico + app.png")


if __name__ == "__main__":
    main()
