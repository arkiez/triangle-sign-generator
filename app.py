# -*- coding: utf-8 -*-
"""
app.py — โปรแกรมสร้างป้ายสามเหลี่ยม (Word) แบบ Mail Merge มีหน้าต่างกรอกข้อมูล
กรอก ชื่อ/ตำแหน่ง/หน่วยงาน -> พรีวิว / บันทึก .docx / เปิดใน Word / ส่งออก PDF / สั่งพิมพ์
รองรับทั้งทีละคน และนำเข้ารายชื่อหลายคนจาก Excel/CSV
"""
import os
import sys
import queue
import tempfile
import threading
import traceback

import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import tkinter.font as tkfont

import merge_engine as me
import word_io as wio
import data_io

APP_VERSION = "1.3.1"
CREATOR = "Powered by Arkie'z K. Khositkhanawut"

# ฟอนต์: ฝัง Kanit ทั้งเนื้อหาและหัวข้อ มากับโปรแกรม โหลดแบบ private
# จึงใช้ได้แม้เครื่องปลายทางไม่ได้ติดตั้งฟอนต์ (ดู _load_private_fonts). ค่าจริงเลือกใน _setup_theme
FONT_DIR = os.path.join(me.BASE, "fonts")
BODY_PREFS = ("Kanit", "Leelawadee UI", "Tahoma")
HEAD_PREFS = ("Kanit", "Leelawadee UI", "Segoe UI", "Tahoma")
FONT_FAMILY = BODY_PREFS[0]   # ฟอนต์เนื้อหา
HEAD_FAMILY = HEAD_PREFS[0]   # ฟอนต์หัวข้อ
UI_FONT = (FONT_FAMILY, 11)
UI_FONT_BOLD = (FONT_FAMILY, 11, "bold")
TITLE_FONT = (HEAD_FAMILY, 16, "bold")
SMALL_FONT = (FONT_FAMILY, 9)


def _load_private_fonts(font_dir=None):
    """โหลดฟอนต์ .ttf ในโฟลเดอร์ fonts/ แบบ private (ใช้ได้โดยไม่ต้องติดตั้งในเครื่อง)
    รองรับ Windows ผ่าน AddFontResourceExW(FR_PRIVATE)"""
    font_dir = font_dir or FONT_DIR
    if sys.platform != "win32" or not os.path.isdir(font_dir):
        return
    import ctypes
    FR_PRIVATE = 0x10
    try:
        gdi32 = ctypes.WinDLL("gdi32")
    except Exception:
        return
    for fn in sorted(os.listdir(font_dir)):
        if fn.lower().endswith((".ttf", ".otf")):
            try:
                gdi32.AddFontResourceExW(os.path.join(font_dir, fn), FR_PRIVATE, 0)
            except Exception:
                pass
PREVIEW_DPI = 70  # ~820px กว้างสำหรับ A4 แนวนอน

# ---- จานสี (Polished Light) ----
COL_BG = "#f5f7fb"          # พื้นหลังหลัก (เทาอมฟ้าอ่อนมาก)
COL_CARD = "#ffffff"        # พื้นผิวการ์ด/พรีวิว
COL_TEXT = "#1e293b"        # ข้อความหลัก
COL_MUTED = "#64748b"       # ข้อความรอง
COL_BORDER = "#e3e8f0"      # เส้นขอบบาง
COL_FIELD = "#ffffff"       # พื้นช่องกรอก
COL_ROW_ALT = "#f8fafc"     # แถวสลับในตาราง
COL_BTN_BORDER = "#cbd5e1"  # ขอบปุ่มรอง

# สีเน้น (accent) — ใช้สีเดียวทั้งโปรแกรม
COL_ACCENT = "#2563eb"
COL_ACCENT_HOVER = "#1d4ed8"
COL_ACCENT_ACTIVE = "#1e40af"
COL_ACCENT_SOFT = "#eff4ff"

COL_HEADER = "#ffffff"      # แถบหัวโทนสว่าง


class App(tk.Tk):
    def __init__(self):
        _load_private_fonts()   # ฝังฟอนต์ก่อนสร้าง Tk root เพื่อให้ Tk เห็นฟอนต์
        super().__init__()
        self.title(f"สร้างป้ายสามเหลี่ยม (Word) — Mail Merge  v{APP_VERSION}")
        self._set_icon()
        self._setup_theme()

        self.template_type = tk.StringVar(value=me.TYPE1)
        self.mode = tk.StringVar(value="single")           # single | batch
        self.batch_output = tk.StringVar(value="combined")  # combined | separate
        self.busy = False
        self._preview_img = None
        self._batch_rows = []
        self._batch_file = None
        self._preview_index = 0
        self._preview_cache = {}  # (template_type, index) -> png path
        self._spin_job = None     # after id ของสปินเนอร์ (None = ไม่หมุน)
        self._spin_angle = 0
        self._spin_text = "กำลังทำงาน…"

        self._build_ui()
        self._install_clipboard()
        self._refresh_visibility()
        self._update_capabilities()
        self._position_window()
        self.protocol("WM_DELETE_WINDOW", self._on_close)

    def _position_window(self):
        """กำหนดขนาด/ตำแหน่งหน้าต่างให้พอดีจอเสมอ (คอลัมน์ซ้ายเลื่อนได้เมื่อเนื้อหายาวเกินจอ)"""
        self.update_idletasks()
        # วัดความสูงเนื้อหาคอลัมน์ซ้ายจากโหมดที่สูงกว่า (batch) แล้วค่อยกลับมาโหมดเดิม
        cur = self.mode.get()
        self.mode.set("batch"); self._refresh_visibility(); self.update_idletasks()
        left_need = self._left_inner.winfo_reqheight()
        self.mode.set(cur); self._refresh_visibility(); self.update_idletasks()

        sw, sh = self.winfo_screenwidth(), self.winfo_screenheight()
        header_h = self.header.winfo_reqheight()
        footer_h = self.footer.winfo_reqheight()
        body = max(left_need, 460)                 # เผื่อพื้นที่พรีวิวด้วย
        desired = header_h + body + footer_h + 40  # เผื่อ padding ต่าง ๆ
        avail = sh - 96                            # เผื่อแถบชื่อ + taskbar
        h = max(min(desired, avail), 520)
        w = min(1200, sw - 40)
        x = max((sw - w) // 2, 0)
        self.minsize(900, 520)
        self.resizable(True, True)                 # ปรับขนาดได้ + ซ้ายเลื่อนได้ -> ไม่มีล้นจอ
        self.geometry(f"{w}x{h}+{x}+8")            # ชิดด้านบนของจอ

    def _set_icon(self):
        """ตั้งไอคอนหน้าต่าง (ใช้ได้ทั้งรันจากซอร์สและจาก .exe)"""
        try:
            ico = os.path.join(me.BASE, "app.ico")
            if os.path.exists(ico):
                self.iconbitmap(ico)
        except Exception:
            pass

    def _pick_font(self, prefs):
        """เลือกฟอนต์ตัวแรกใน prefs ที่ติดตั้งจริงในเครื่อง (กันเครื่องที่ไม่มี Leelawadee UI)"""
        try:
            available = set(tkfont.families(self))
        except tk.TclError:
            available = set()
        for fam in prefs:
            if fam in available:
                return fam
        return prefs[-1]

    def _setup_theme(self):
        """ตั้งฟอนต์ฝัง (เนื้อหา=TH SarabunPSK, หัวข้อ=Kanit) + จานสี (ใช้ธีม clam ปรับสีได้เต็มที่)"""
        # เลือกฟอนต์ที่โหลด/ติดตั้งได้จริง แล้วอัปเดตค่าฟอนต์ให้ทั้งโปรแกรมใช้ (เรียกก่อน _build_ui)
        global FONT_FAMILY, HEAD_FAMILY, UI_FONT, UI_FONT_BOLD, TITLE_FONT, SMALL_FONT
        FONT_FAMILY = self._pick_font(BODY_PREFS)
        HEAD_FAMILY = self._pick_font(HEAD_PREFS)
        UI_FONT = (FONT_FAMILY, 11)
        UI_FONT_BOLD = (FONT_FAMILY, 11, "bold")
        TITLE_FONT = (HEAD_FAMILY, 16, "bold")
        SMALL_FONT = (FONT_FAMILY, 9)

        # ฟอนต์เริ่มต้นของ Tk (กล่องข้อความ/เมนู) ใช้ฟอนต์เนื้อหา
        for fname in ("TkDefaultFont", "TkTextFont", "TkMenuFont",
                      "TkHeadingFont", "TkTooltipFont", "TkIconFont"):
            try:
                tkfont.nametofont(fname).configure(family=FONT_FAMILY, size=10)
            except tk.TclError:
                pass

        style = ttk.Style(self)
        try:
            style.theme_use("clam")
        except tk.TclError:
            pass

        self.configure(bg=COL_BG)
        # ค่าตั้งต้น: พื้นผิวเป็นการ์ดสีขาว (ส่วนโครงสร้างที่ต้องเป็นพื้นเทาใช้สไตล์ Surface.*)
        style.configure(".", font=(FONT_FAMILY, 11), background=COL_CARD, foreground=COL_TEXT)
        style.configure("TFrame", background=COL_CARD)
        style.configure("TLabel", background=COL_CARD, foreground=COL_TEXT)
        # พื้นผิวโครงสร้าง (พื้นเทาอ่อน) สำหรับนอกการ์ด
        style.configure("Surface.TFrame", background=COL_BG)
        style.configure("Surface.TLabel", background=COL_BG, foreground=COL_MUTED)

        # การ์ด = กรอบขาว เส้นขอบบาง หัวข้อใช้ฟอนต์หัวข้อ (Kanit) สีเน้น
        style.configure("TLabelframe", background=COL_CARD, bordercolor=COL_BORDER,
                        relief="solid", borderwidth=1)
        style.configure("TLabelframe.Label", background=COL_CARD, foreground=COL_ACCENT,
                        font=(HEAD_FAMILY, 12, "bold"))

        # เรดิโอ: ชี้แล้วตัวอักษรเป็นสีเน้น จุดเลือกเป็นสีเน้น
        style.configure("TRadiobutton", background=COL_CARD, foreground=COL_TEXT,
                        font=(FONT_FAMILY, 11))
        style.map("TRadiobutton",
                  background=[("active", COL_CARD)],
                  foreground=[("active", COL_ACCENT)],
                  indicatorcolor=[("selected", COL_ACCENT), ("pressed", COL_ACCENT)])

        # ช่องกรอก: พื้นขาว ขอบบาง โฟกัสแล้วขอบเป็นสีเน้น
        style.configure("TEntry", fieldbackground=COL_FIELD, foreground=COL_TEXT,
                        bordercolor=COL_BTN_BORDER, relief="solid", borderwidth=1,
                        padding=(6, 5))
        style.map("TEntry",
                  bordercolor=[("focus", COL_ACCENT)],
                  lightcolor=[("focus", COL_ACCENT)],
                  darkcolor=[("focus", COL_ACCENT)])
        style.configure("TCombobox", fieldbackground=COL_FIELD)

        # ตาราง: แถวสูงขึ้น หัวตารางใช้ฟอนต์หัวข้อ (Kanit) เลือกแถวเป็นสีเน้นอ่อน
        style.configure("Treeview", font=(FONT_FAMILY, 11), rowheight=27,
                        fieldbackground=COL_CARD, background=COL_CARD,
                        foreground=COL_TEXT, borderwidth=0)
        style.configure("Treeview.Heading", font=(HEAD_FAMILY, 11, "bold"),
                        background=COL_BG, foreground=COL_TEXT,
                        relief="flat", padding=(8, 7))
        style.map("Treeview.Heading", background=[("active", COL_ACCENT_SOFT)])
        style.map("Treeview",
                  background=[("selected", COL_ACCENT_SOFT)],
                  foreground=[("selected", COL_ACCENT)])

        # แถบเลื่อน: โทนกลาง เรียบ
        for sb in ("Vertical.TScrollbar", "Horizontal.TScrollbar"):
            style.configure(sb, background=COL_BORDER, troughcolor=COL_BG,
                            bordercolor=COL_BG, arrowcolor=COL_MUTED, relief="flat")
            style.map(sb, background=[("active", COL_BTN_BORDER)])

        # ปุ่มรอง (ค่าเริ่มต้น): ขาว ขอบบาง ชี้แล้วเป็นฟ้าอ่อน + ตัวอักษรสีเน้น
        for sname, pad in (("TButton", (10, 8)), ("Secondary.TButton", (10, 9))):
            style.configure(sname, font=(FONT_FAMILY, 11, "bold"), padding=pad,
                            background=COL_CARD, foreground=COL_TEXT,
                            bordercolor=COL_BTN_BORDER, focuscolor=COL_CARD,
                            relief="solid", borderwidth=1)
            style.map(sname,
                      background=[("active", COL_ACCENT_SOFT), ("pressed", COL_ACCENT_SOFT),
                                  ("disabled", COL_BG)],
                      foreground=[("active", COL_ACCENT), ("disabled", "#9aa6b8")],
                      bordercolor=[("active", COL_ACCENT), ("disabled", COL_BORDER)])

        # ปุ่มเลื่อนพรีวิว ◀ ▶ : กะทัดรัด สีเน้น
        style.configure("Nav.TButton", font=(FONT_FAMILY, 11, "bold"), padding=(6, 2),
                        background=COL_CARD, foreground=COL_ACCENT,
                        bordercolor=COL_BORDER, focuscolor=COL_CARD,
                        relief="solid", borderwidth=1)
        style.map("Nav.TButton",
                  background=[("active", COL_ACCENT_SOFT)],
                  foreground=[("active", COL_ACCENT_ACTIVE), ("disabled", "#9aa6b8")])

        # ปุ่มหลัก (CTA): ทึบสีเน้น
        style.configure("Primary.TButton", font=(FONT_FAMILY, 12, "bold"), padding=(10, 10),
                        background=COL_ACCENT, foreground="white",
                        bordercolor=COL_ACCENT, focuscolor=COL_ACCENT, relief="flat")
        style.map("Primary.TButton",
                  background=[("active", COL_ACCENT_HOVER), ("pressed", COL_ACCENT_ACTIVE),
                              ("disabled", "#c7d2e4")],
                  foreground=[("disabled", "#eef2f8")])

    # --------------------------------------------------- พื้นที่เลื่อน (ซ้าย)
    def _make_scroll_column(self, parent, width):
        """สร้างคอลัมน์ที่เลื่อนแนวตั้งได้ คืน (container, inner)
        ใส่ widget ลงใน inner ตามปกติ; ถ้าเนื้อหายาวเกินจอจะมีแถบเลื่อนให้เอง"""
        container = ttk.Frame(parent, style="Surface.TFrame")
        container.rowconfigure(0, weight=1)
        container.columnconfigure(0, weight=1)

        canvas = tk.Canvas(container, bg=COL_BG, highlightthickness=0, width=width)
        canvas.grid(row=0, column=0, sticky="nsew")
        sb = ttk.Scrollbar(container, orient="vertical", command=canvas.yview)
        canvas.configure(yscrollcommand=sb.set)

        inner = ttk.Frame(canvas, style="Surface.TFrame")
        win = canvas.create_window((0, 0), window=inner, anchor="nw")

        def _sync(_=None):
            canvas.configure(scrollregion=canvas.bbox("all"))
            need = inner.winfo_reqheight() > canvas.winfo_height()
            if need:
                sb.grid(row=0, column=1, sticky="ns")
            else:
                sb.grid_remove()
                canvas.yview_moveto(0)

        def _on_canvas(e):
            canvas.itemconfigure(win, width=e.width)  # ให้ inner กว้างเท่าผืนผ้าใบ
            _sync()

        inner.bind("<Configure>", _sync)
        canvas.bind("<Configure>", _on_canvas)

        def _wheel(e):
            # ถ้าตัวชี้อยู่บนตารางรายชื่อ ปล่อยให้ตารางเลื่อนเอง
            w = self.winfo_containing(e.x_root, e.y_root)
            tree = getattr(self, "tree", None)
            p = w
            while p is not None:
                if p is tree:
                    return
                p = getattr(p, "master", None)
            if inner.winfo_reqheight() > canvas.winfo_height():
                canvas.yview_scroll(int(-e.delta / 120), "units")

        canvas.bind("<Enter>", lambda e: canvas.bind_all("<MouseWheel>", _wheel))
        canvas.bind("<Leave>", lambda e: canvas.unbind_all("<MouseWheel>"))
        return container, inner

    # ------------------------------------------------------- คลิปบอร์ด/คีย์
    def _install_clipboard(self):
        """ผูกคีย์ Ctrl+C/V/X/A กับช่องกรอกทุกช่อง โดยอิง keycode (physical key)
        เพื่อให้คัดลอก/วางได้แม้สลับแป้นเป็นภาษาไทย (เป็นบั๊กยอดฮิตของ Tkinter)"""
        self.bind_class("TEntry", "<Control-KeyPress>", self._entry_clip, add="+")

    def _entry_clip(self, event):
        """จัดการคัดลอก/วาง/ตัด/เลือกทั้งหมด ในช่องกรอก (ไม่พึ่ง keysym ของแป้น)"""
        w = event.widget
        kc = event.keycode
        try:
            if kc == 67:        # C — คัดลอก
                if w.selection_present():
                    self.clipboard_clear()
                    self.clipboard_append(w.get()[w.index("sel.first"):w.index("sel.last")])
                return "break"
            if kc == 88:        # X — ตัด
                if w.selection_present():
                    self.clipboard_clear()
                    self.clipboard_append(w.get()[w.index("sel.first"):w.index("sel.last")])
                    w.delete("sel.first", "sel.last")
                return "break"
            if kc == 86:        # V — วาง
                try:
                    text = self.clipboard_get()
                except tk.TclError:
                    text = ""
                if text:
                    if w.selection_present():
                        w.delete("sel.first", "sel.last")
                    # วางในช่องเดียว: ตัดบรรทัด/แท็บออกกันเพี้ยน
                    w.insert("insert", text.replace("\r", " ").replace("\n", " ").replace("\t", " "))
                return "break"
            if kc == 65:        # A — เลือกทั้งหมด
                w.select_range(0, "end")
                w.icursor("end")
                return "break"
        except tk.TclError:
            pass
        return None

    def _tree_key(self, event):
        """คีย์ลัดบนตารางรายชื่อ: Ctrl+C คัดลอก, Ctrl+V วาง, Ctrl+X ตัด, Ctrl+A เลือกทั้งหมด"""
        kc = event.keycode
        if kc == 67:
            self._copy_tree(); return "break"
        if kc == 86:
            self._paste_tree(); return "break"
        if kc == 88:
            self._copy_tree(); self._delete_selected(); return "break"
        if kc == 65:
            self.tree.selection_set(self.tree.get_children()); return "break"
        return None

    def _copy_tree(self):
        """คัดลอกแถวที่เลือกลงคลิปบอร์ดเป็น TSV (วางต่อใน Excel ได้)"""
        sel = self.tree.selection()
        if not sel:
            return
        lines = ["\t".join(str(x) for x in self.tree.item(iid, "values")) for iid in sel]
        self.clipboard_clear()
        self.clipboard_append("\n".join(lines))
        self._set_status(f"คัดลอก {len(sel)} รายการแล้ว", "#080")

    def _paste_tree(self):
        """วางข้อความจากคลิปบอร์ดเป็นแถวใหม่ (รองรับหลายแถวจาก Excel: คั่นด้วยแท็บ/จุลภาค)"""
        if self.mode.get() != "batch":
            return
        try:
            data = self.clipboard_get()
        except tk.TclError:
            return
        rows = []
        for line in data.replace("\r\n", "\n").replace("\r", "\n").split("\n"):
            if not line.strip():
                continue
            parts = line.split("\t") if "\t" in line else line.split(",")
            parts = [p.strip() for p in parts]
            name = parts[0] if parts else ""
            if not name:
                continue
            rows.append((name,
                         parts[1] if len(parts) > 1 else "",
                         parts[2] if len(parts) > 2 else ""))
        if not rows:
            return
        sel = self.tree.selection()
        base = self.tree.index(sel[-1]) + 1 if sel else None  # วางต่อจากแถวที่เลือก
        last = None
        for i, r in enumerate(rows):
            last = self.tree.insert("", "end" if base is None else base + i, values=r)
        self._sync_rows_from_tree()
        if not self.navbar.winfo_ismapped():
            self.navbar.pack(fill="x", pady=(0, 6), before=self.canvas)
        if last:
            self.tree.selection_set(last)
            self.tree.see(last)
        self._set_status(f"วาง {len(rows)} รายการแล้ว", "#080")

    # ------------------------------------------------------------------ UI
    def _build_ui(self):
        # ---- แถบหัวโปรแกรม (สว่าง + เส้นคั่นบาง) ----
        header = tk.Frame(self, bg=COL_HEADER)
        self.header = header
        header.pack(fill="x", side="top")
        tk.Label(header, text="▲", bg=COL_HEADER, fg=COL_ACCENT,
                 font=(HEAD_FAMILY, 18, "bold")).pack(side="left", padx=(18, 8), pady=12)
        tk.Label(header, text="สร้างป้ายสามเหลี่ยม", bg=COL_HEADER, fg=COL_TEXT,
                 font=(HEAD_FAMILY, 18, "bold")).pack(side="left", pady=12)
        tk.Label(header, text=f"Mail Merge · v{APP_VERSION}", bg=COL_HEADER, fg=COL_MUTED,
                 font=(FONT_FAMILY, 10)).pack(side="right", padx=18)
        tk.Frame(self, bg=COL_BORDER, height=1).pack(fill="x", side="top")  # เส้นคั่นใต้หัว

        root = ttk.Frame(self, padding=12, style="Surface.TFrame")
        root.pack(fill="both", expand=True)
        root.columnconfigure(0, weight=0, minsize=430)
        root.columnconfigure(1, weight=1)
        root.rowconfigure(0, weight=1)

        # คอลัมน์ซ้ายห่อด้วยพื้นที่เลื่อนแนวตั้ง (กันเนื้อหาล้นจอบนหน้าจอเล็ก)
        left_container, left = self._make_scroll_column(root, width=430)
        left_container.grid(row=0, column=0, sticky="nsew", padx=(0, 12))
        self._left_inner = left
        right = ttk.Frame(root, style="Surface.TFrame")
        right.grid(row=0, column=1, sticky="nsew")
        self.right = right

        # ---- ส่วนท้าย: เวอร์ชัน + ผู้สร้าง ----
        footer = ttk.Frame(root, style="Surface.TFrame")
        self.footer = footer
        footer.grid(row=1, column=0, columnspan=2, sticky="ew", pady=(10, 0))
        # ขวา: เวอร์ชัน + ผู้สร้าง | ซ้าย: สถานะการทำงาน (status bar)
        ttk.Label(footer, text=f"v{APP_VERSION} · {CREATOR}",
                  style="Surface.TLabel", font=SMALL_FONT).pack(side="right")
        self.status = ttk.Label(footer, text="พร้อมใช้งาน", style="Surface.TLabel", font=UI_FONT_BOLD)
        self.status.pack(side="left")

        # ---- ซ้าย: ตัวควบคุม ----
        # แบบป้าย
        tf = ttk.LabelFrame(left, text=" แบบป้าย ", padding=12)
        tf.pack(fill="x")
        ttk.Radiobutton(tf, text=me.TYPE_LABELS[me.TYPE1], value=me.TYPE1,
                        variable=self.template_type, command=self._refresh_visibility).pack(anchor="w", pady=2)
        ttk.Radiobutton(tf, text=me.TYPE_LABELS[me.TYPE2], value=me.TYPE2,
                        variable=self.template_type, command=self._refresh_visibility).pack(anchor="w", pady=2)

        # โหมด
        mf = ttk.LabelFrame(left, text=" โหมด ", padding=12)
        mf.pack(fill="x", pady=(10, 0))
        self.mf = mf
        ttk.Radiobutton(mf, text="ทีละคน", value="single",
                        variable=self.mode, command=self._refresh_visibility).pack(side="left", padx=(0, 16))
        ttk.Radiobutton(mf, text="นำเข้ารายชื่อ (Excel/CSV)", value="batch",
                        variable=self.mode, command=self._refresh_visibility).pack(side="left")

        # ---- ฟอร์มทีละคน ----
        self.single_frame = ttk.LabelFrame(left, text=" ข้อมูล ", padding=12)
        self.single_frame.pack(fill="x", pady=(10, 0))

        ttk.Label(self.single_frame, text="ชื่อ", font=UI_FONT_BOLD).grid(row=0, column=0, sticky="w", pady=4)
        self.e_name = ttk.Entry(self.single_frame, font=UI_FONT, width=34)
        self.e_name.grid(row=0, column=1, sticky="ew", pady=4)
        # พิมพ์ชื่อ -> เปิด/ปิดปุ่มสร้างตามว่ามีข้อมูลหรือยัง
        self.e_name.bind("<KeyRelease>", lambda e: self._refresh_buttons())

        self.lbl_pos = ttk.Label(self.single_frame, text="ตำแหน่ง", font=UI_FONT_BOLD)
        self.lbl_pos.grid(row=1, column=0, sticky="w", pady=4)
        self.e_pos = ttk.Entry(self.single_frame, font=UI_FONT, width=34)
        self.e_pos.grid(row=1, column=1, sticky="ew", pady=4)

        self.lbl_org = ttk.Label(self.single_frame, text="หน่วยงาน", font=UI_FONT_BOLD)
        self.lbl_org.grid(row=2, column=0, sticky="w", pady=4)
        self.e_org = ttk.Entry(self.single_frame, font=UI_FONT, width=34)
        self.e_org.grid(row=2, column=1, sticky="ew", pady=4)
        self.single_frame.columnconfigure(1, weight=1)

        # ---- ฟอร์ม batch ----
        self.batch_frame = ttk.LabelFrame(left, text=" รายชื่อ ", padding=12)
        # (pack/forget ใน _refresh_visibility)
        topb = ttk.Frame(self.batch_frame)
        topb.pack(fill="x")
        ttk.Button(topb, text="📂 Choose Excel/CSV…", command=self._choose_list).pack(side="left")
        self.lbl_batch = ttk.Label(topb, text="ยังไม่ได้เลือกไฟล์", font=UI_FONT)
        self.lbl_batch.pack(side="left", padx=10)

        tk.Label(self.batch_frame, bg=COL_CARD, fg=COL_MUTED, font=SMALL_FONT, justify="left",
                 text="ดับเบิลคลิกเพื่อแก้ในช่อง · Ctrl+C คัดลอก · Ctrl+V วางจาก Excel · Del ลบแถว"
                 ).pack(anchor="w", pady=(6, 0))
        cols = ("name", "position", "org")
        treewrap = ttk.Frame(self.batch_frame)
        treewrap.pack(fill="both", expand=True, pady=(2, 6))
        self.tree = ttk.Treeview(treewrap, columns=cols, show="headings", height=5)
        for c, t, w in (("name", "ชื่อ", 150), ("position", "ตำแหน่ง", 110), ("org", "หน่วยงาน", 120)):
            self.tree.heading(c, text=t)
            self.tree.column(c, width=w, anchor="w")
        self.tree.tag_configure("odd", background=COL_CARD)
        self.tree.tag_configure("even", background=COL_ROW_ALT)
        vsb = ttk.Scrollbar(treewrap, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=vsb.set)
        self.tree.pack(side="left", fill="both", expand=True)
        vsb.pack(side="right", fill="y")
        self.tree.bind("<Double-1>", self._edit_cell)
        # คัดลอก/วาง/ลบ ในตาราง (รองรับทั้งคีย์บอร์ดไทยและอังกฤษผ่าน keycode)
        self.tree.bind("<Control-KeyPress>", self._tree_key)
        self.tree.bind("<Delete>", lambda e: (self._delete_selected(), "break")[1])

        # ปุ่มเพิ่ม/ลบรายชื่อ
        editbar = ttk.Frame(self.batch_frame)
        editbar.pack(fill="x", pady=(0, 6))
        ttk.Button(editbar, text="➕ เพิ่มคน", command=self._add_person).pack(side="left")
        ttk.Button(editbar, text="🗑 ลบที่เลือก", command=self._delete_selected).pack(side="left", padx=6)
        ttk.Button(editbar, text="🧹 ล้างทั้งหมด", command=self._clear_all).pack(side="right")

        ob = ttk.Frame(self.batch_frame)
        ob.pack(fill="x")
        ttk.Radiobutton(ob, text="รวมไฟล์เดียว", value="combined",
                        variable=self.batch_output).pack(side="left", padx=(0, 14))
        ttk.Radiobutton(ob, text="แยกไฟล์ต่อคน", value="separate",
                        variable=self.batch_output).pack(side="left")

        # ---- ปุ่มคำสั่ง (จัดเป็นกริด 2 คอลัมน์ ให้ประหยัดความสูง) ----
        bf = ttk.Frame(left, style="Surface.TFrame")
        bf.pack(fill="x", pady=(12, 0))
        bf.columnconfigure(0, weight=1)
        bf.columnconfigure(1, weight=1)
        self.btn_preview = ttk.Button(bf, text="🔍  Preview", style="Primary.TButton",
                                      command=self._on_preview)
        self.btn_save = ttk.Button(bf, text="💾  Save .docx", style="Secondary.TButton",
                                   command=self._on_save)
        self.btn_open = ttk.Button(bf, text="📝  Open in Word", style="Secondary.TButton",
                                   command=self._on_open)
        self.btn_pdf = ttk.Button(bf, text="📄  Export PDF", style="Secondary.TButton",
                                  command=self._on_pdf)
        self.btn_print = ttk.Button(bf, text="🖨  Print", style="Secondary.TButton",
                                    command=self._on_print)
        self.btn_preview.grid(row=0, column=0, columnspan=2, sticky="ew", pady=(0, 6))
        self.btn_save.grid(row=1, column=0, sticky="ew", padx=(0, 4), pady=3)
        self.btn_open.grid(row=1, column=1, sticky="ew", padx=(4, 0), pady=3)
        self.btn_pdf.grid(row=2, column=0, sticky="ew", padx=(0, 4), pady=3)
        self.btn_print.grid(row=2, column=1, sticky="ew", padx=(4, 0), pady=3)

        # ---- ขวา: พรีวิว ----
        pv = ttk.LabelFrame(right, text=" พรีวิว ", padding=12)
        pv.pack(fill="both", expand=True)

        # แถบเลื่อนดู (เฉพาะโหมดนำเข้ารายชื่อ)
        self.navbar = ttk.Frame(pv)
        self.btn_prev = ttk.Button(self.navbar, text="◀", width=3, style="Nav.TButton",
                                   command=self._preview_prev)
        self.btn_prev.pack(side="left")
        self.lbl_nav = ttk.Label(self.navbar, text="", font=UI_FONT_BOLD, foreground=COL_ACCENT)
        self.lbl_nav.pack(side="left", padx=10)
        self.btn_next = ttk.Button(self.navbar, text="▶", width=3, style="Nav.TButton",
                                   command=self._preview_next)
        self.btn_next.pack(side="left")

        self.canvas = tk.Canvas(pv, background=COL_CARD, highlightthickness=0)
        self.canvas.pack(fill="both", expand=True)
        self.canvas.create_text(
            14, 14, anchor="nw", fill=COL_MUTED, font=UI_FONT,
            text="กรอกข้อมูลแล้วกดปุ่ม \"Preview\" เพื่อดูภาพป้ายก่อนบันทึก",
        )

        # ปุ่มที่ต้องปิดระหว่างทำงาน
        self._busy_buttons = [self.btn_preview, self.btn_save, self.btn_open,
                              self.btn_pdf, self.btn_print, self.btn_prev, self.btn_next]

    def _refresh_visibility(self):
        is_t1 = self.template_type.get() == me.TYPE1
        for w in (self.lbl_pos, self.e_pos, self.lbl_org, self.e_org):
            if is_t1:
                w.grid()
            else:
                w.grid_remove()

        if self.mode.get() == "single":
            self.batch_frame.pack_forget()
            self.single_frame.pack(fill="x", pady=(10, 0), after=self.mf)
        else:
            self.single_frame.pack_forget()
            self.batch_frame.pack(fill="both", expand=False, pady=(10, 0), after=self.mf)

        # เปลี่ยนแบบป้าย/โหมด -> แคชพรีวิวเดิมใช้ไม่ได้
        self._preview_cache.clear()
        # แสดง/ซ่อนแถบเลื่อนดู (เฉพาะโหมดนำเข้ารายชื่อที่มีข้อมูลแล้ว)
        if self.mode.get() == "batch" and self._batch_rows:
            self.navbar.pack(fill="x", pady=(0, 6), before=self.canvas)
        else:
            self.navbar.pack_forget()
        self._update_nav_label()
        self._refresh_buttons()

    def _update_capabilities(self):
        """แจ้งสถานะเมื่อขาด Word/PyMuPDF แล้วรีเฟรชสถานะปุ่ม"""
        ok_word = wio.pywin32_available()
        ok_fitz = wio.fitz_available()
        if not ok_word:
            self._set_status("ไม่พบ pywin32/Word — ใช้ได้เฉพาะบันทึก .docx", "#a00")
        elif not ok_fitz:
            self._set_status("ไม่พบ PyMuPDF — พรีวิวรูปปิดอยู่ (ฟีเจอร์อื่นใช้ได้)", "#a60")
        self._refresh_buttons()

    def _has_input(self):
        """ฟอร์มมีข้อมูลพอจะสร้างป้ายไหม (กันสร้างจากฟอร์มว่าง)"""
        if self.mode.get() == "single":
            return bool(self.e_name.get().strip())
        return any((r.get("name") or "").strip() for r in self._batch_rows)

    def _refresh_buttons(self):
        """เปิด/ปิดปุ่มตามความพร้อม: ต้องมีข้อมูล + มี Word/PyMuPDF ตามที่ปุ่มนั้นต้องใช้
        (ฟอร์มว่าง = ปุ่มสร้างทั้งหมดถูกปิด จึงกด generate ไม่ได้)"""
        if self.busy:
            return
        ok_word = wio.pywin32_available()
        ok_fitz = wio.fitz_available()
        has = self._has_input()

        def setb(btn, enabled):
            btn.state(["!disabled"] if enabled else ["disabled"])

        setb(self.btn_save, has)                       # บันทึก .docx ไม่ต้องใช้ Word
        setb(self.btn_open, has and ok_word)
        setb(self.btn_pdf, has and ok_word)
        setb(self.btn_print, has and ok_word)
        setb(self.btn_preview, has and ok_word and ok_fitz)
        nav_ok = (self.mode.get() == "batch" and bool(self._batch_rows)
                  and ok_word and ok_fitz)
        setb(self.btn_prev, nav_ok)
        setb(self.btn_next, nav_ok)

    # ------------------------------------------------------------- helpers
    def _set_status(self, text, color="#246"):
        self.status.configure(text=text, foreground=color)

    def _snapshot(self):
        """อ่านค่าจากหน้าจอบนเธรดหลัก -> เป็นข้อมูลล้วน (ส่งให้ worker ใช้ได้อย่างปลอดภัย)
        Tkinter เรียกข้ามเธรดไม่ได้ จึง snapshot ก่อนแล้วค่อยส่งงานเข้า worker"""
        return {
            "ttype": self.template_type.get(),
            "mode": self.mode.get(),
            "batch_output": self.batch_output.get(),
            "single": {
                "name": self.e_name.get().strip(),
                "position": self.e_pos.get().strip(),
                "org": self.e_org.get().strip(),
            },
            "rows": list(self._batch_rows),
            "preview_index": self._preview_index,
        }

    def _gen_docx(self, job, out_path, first_only=False):
        """สร้างไฟล์ .docx จาก snapshot (งานล้วน ไม่แตะ Tk -> เรียกบน worker ได้)
        คืนค่า path (หรือ list สำหรับ separate)"""
        ttype = job["ttype"]
        if job["mode"] == "single":
            d = job["single"]
            return me.generate_one(ttype, d["name"], d["position"], d["org"], out_path=out_path)
        rows = job["rows"]
        if first_only:
            r = rows[0]
            return me.generate_one(ttype, r["name"], r.get("position", ""), r.get("org", ""), out_path=out_path)
        if job["batch_output"] == "separate":
            return me.generate_batch(ttype, rows, out_path, separate=True)
        return me.generate_batch(ttype, rows, out_path, separate=False)[0]

    def _default_name(self):
        if self.mode.get() == "single":
            return me.sanitize_filename(self.e_name.get())
        if self._batch_rows:
            return "ป้ายสามเหลี่ยม_หลายคน"
        return "ป้ายสามเหลี่ยม"

    # ------------------------------------------------------- async runner
    # Tkinter เรียกได้จากเธรดหลักเท่านั้น (เรียก after()/dialog ข้ามเธรดจะ error
    # ทำให้ worker ตายเงียบ ๆ และโปรแกรมค้างที่ "กำลังทำงาน…")
    # จึงให้ worker ทำเฉพาะงานล้วน แล้วส่งผลกลับผ่านคิว โดยเธรดหลักเป็นผู้ poll
    def _do(self, action):
        if self.busy:
            return
        self.busy = True
        for b in self._busy_buttons:
            b.state(["disabled"])
        self._set_status("กำลังทำงาน…", "#a60")

        q = queue.Queue()

        def worker():
            try:
                q.put(("ok", action()))
            except Exception as e:
                q.put(("err", (e, traceback.format_exc())))

        threading.Thread(target=worker, daemon=True).start()
        self._poll_job(q)

    def _poll_job(self, q):
        """ตรวจคิวผลลัพธ์บนเธรดหลัก (re-schedule ด้วย after จากเธรดหลัก = ปลอดภัย)"""
        try:
            kind, payload = q.get_nowait()
        except queue.Empty:
            self.after(80, lambda: self._poll_job(q))
            return
        if kind == "ok":
            self._done(payload, None)
        else:
            self._done(None, payload)

    def _done(self, result, err):
        self.busy = False
        self._stop_spinner()    # หยุดสปินเนอร์ก่อนแสดงผล/แจ้ง error
        self._refresh_buttons()
        if err:
            e, tb = err
            self._set_status(f"ผิดพลาด: {e}", "#a00")
            messagebox.showerror("เกิดข้อผิดพลาด", str(e))
            return
        if not (isinstance(result, tuple) and result):
            return
        tag = result[0]
        if tag == "preview":
            _, png, msg = result
            self._show_preview(png)
            self._set_status(msg, "#080")
        elif tag == "reveal":
            _, folder, msg = result
            self._reveal(folder)
            self._set_status(msg, "#080")
        elif tag == "status":
            self._set_status(result[1], "#080")

    # ---------------------------------------------------------- preview
    def _on_preview(self):
        if self.busy:
            return
        try:
            self._precheck()
        except Exception as e:
            messagebox.showinfo("ยังพรีวิวไม่ได้", str(e))
            return
        job = self._snapshot()
        self._start_spinner()   # โชว์สปินเนอร์กลางพรีวิวระหว่างสร้างภาพ
        self._do(lambda: self._preview_worker(job))

    def _preview_worker(self, job):
        """งานพรีวิวล้วน (worker): สร้าง docx -> เรนเดอร์ PNG แล้วคืน path ให้เธรดหลักแสดง"""
        ttype = job["ttype"]
        if job["mode"] == "batch":
            idx = min(job["preview_index"], len(job["rows"]) - 1)
            key = (ttype, idx)
            png = self._preview_cache.get(key)
            if not (png and os.path.exists(png)):  # ยังไม่มีในแคช -> สร้างใหม่
                r = job["rows"][idx]
                tmp = os.path.join(tempfile.gettempdir(), f"_tent_prev_{idx}.docx")
                docx = me.generate_one(ttype, r["name"], r.get("position", ""), r.get("org", ""), out_path=tmp)
                png = os.path.join(tempfile.gettempdir(), f"_tent_prev_{idx}.png")
                png = wio.render_preview_png(docx, png, dpi=PREVIEW_DPI)
                self._preview_cache[key] = png
            return ("preview", png, f"พรีวิว คนที่ {idx + 1}/{len(job['rows'])}")

        # โหมดทีละคน
        d = job["single"]
        tmp = os.path.join(tempfile.gettempdir(), "_tent_preview_src.docx")
        docx = me.generate_one(ttype, d["name"], d["position"], d["org"], out_path=tmp)
        png = wio.render_preview_png(docx, dpi=PREVIEW_DPI)
        return ("preview", png, "พรีวิวเรียบร้อย")

    def _preview_prev(self):
        if self.busy or self.mode.get() != "batch" or not self._batch_rows:
            return
        self._preview_index = (self._preview_index - 1) % len(self._batch_rows)
        self._update_nav_label()
        self._on_preview()

    def _preview_next(self):
        if self.busy or self.mode.get() != "batch" or not self._batch_rows:
            return
        self._preview_index = (self._preview_index + 1) % len(self._batch_rows)
        self._update_nav_label()
        self._on_preview()

    def _update_nav_label(self):
        if self.mode.get() == "batch" and self._batch_rows:
            n = len(self._batch_rows)
            i = min(self._preview_index, n - 1)
            self.lbl_nav.configure(text=f"คนที่ {i + 1} / {n}  ·  {self._batch_rows[i]['name']}")
        else:
            self.lbl_nav.configure(text="")

    # ----------------------------------------------- สปินเนอร์ "กำลังทำงาน"
    def _start_spinner(self, text="กำลังทำงาน…"):
        """แสดงสปินเนอร์หมุนกลางพื้นที่พรีวิว ระหว่างรองานในเธรดเบื้องหลัง
        วาดด้วย create_arc (เวกเตอร์) จึงหมุนได้ลื่นและไม่ต้องพึ่งฟอนต์ไอคอน"""
        self._stop_spinner()
        self.canvas.delete("all")
        self._spin_angle = 0
        self._spin_text = text
        self._tick_spinner()

    def _tick_spinner(self):
        """วาดสปินเนอร์ 1 เฟรมแล้วนัดวาดเฟรมถัดไป (หมุนทีละ 30°)"""
        self.canvas.delete("spinner")
        cw = self.canvas.winfo_width() or 700
        ch = self.canvas.winfo_height() or 500
        r = 24
        box = (cw // 2 - r, ch // 2 - r, cw // 2 + r, ch // 2 + r)
        self.canvas.create_oval(*box, outline=COL_BORDER, width=5, tags="spinner")  # วงราง
        self.canvas.create_arc(*box, start=self._spin_angle, extent=100, style="arc",
                               outline=COL_ACCENT, width=5, tags="spinner")          # ส่วนหมุน
        self.canvas.create_text(cw // 2, ch // 2 + r + 24, text=self._spin_text,
                                fill=COL_MUTED, font=UI_FONT, tags="spinner")
        self._spin_angle = (self._spin_angle - 30) % 360
        self._spin_job = self.after(80, self._tick_spinner)

    def _stop_spinner(self):
        """หยุดสปินเนอร์ + ลบออกจากผืนผ้าใบ (ปลอดภัยแม้ไม่ได้หมุนอยู่)"""
        if self._spin_job is not None:
            try:
                self.after_cancel(self._spin_job)
            except Exception:
                pass
            self._spin_job = None
        self.canvas.delete("spinner")

    def _show_preview(self, png_path):
        try:
            img = tk.PhotoImage(file=png_path)
        except Exception as e:
            self._set_status(f"แสดงรูปไม่ได้: {e}", "#a00")
            return
        self._preview_img = img  # กัน garbage collect
        self.canvas.delete("all")
        cw = self.canvas.winfo_width() or 700
        ch = self.canvas.winfo_height() or 500
        # ย่อแบบจำนวนเต็มให้พอดีกรอบ
        factor = 1
        while (img.width() // factor) > cw or (img.height() // factor) > ch:
            factor += 1
        shown = img.subsample(factor) if factor > 1 else img
        self._preview_img_shown = shown
        self.canvas.create_image(cw // 2, ch // 2, image=shown, anchor="center")
        self._update_nav_label()

    # ---------------------------------------------------- save / open / pdf
    def _on_save(self):
        if self.busy:
            return
        try:
            self._precheck()
        except Exception as e:
            messagebox.showinfo("ยังบันทึกไม่ได้", str(e))
            return
        job = self._snapshot()
        name = self._default_name()
        if job["mode"] == "batch" and job["batch_output"] == "separate":
            folder = filedialog.askdirectory(title="เลือกโฟลเดอร์สำหรับบันทึกไฟล์แยกต่อคน",
                                             initialdir=me.OUTPUT_DIR)
            if not folder:
                self._set_status("ยกเลิกการบันทึก", "#a60")
                return
            self._do(lambda: self._save_sep(job, folder))
            return
        out = filedialog.asksaveasfilename(
            title="บันทึกไฟล์ Word", defaultextension=".docx",
            filetypes=[("Word Document", "*.docx")],
            initialdir=me.OUTPUT_DIR, initialfile=name + ".docx",
        )
        if not out:
            self._set_status("ยกเลิกการบันทึก", "#a60")
            return
        self._do(lambda: self._save_one(job, out))

    def _save_sep(self, job, folder):
        paths = self._gen_docx(job, folder)
        return ("reveal", folder, f"บันทึกแล้ว {len(paths)} ไฟล์ที่ {folder}")

    def _save_one(self, job, out):
        docx = self._gen_docx(job, out)
        return ("reveal", os.path.dirname(docx), f"บันทึกแล้ว: {docx}")

    def _on_open(self):
        if self.busy:
            return
        try:
            self._precheck()
        except Exception as e:
            messagebox.showinfo("ยังเปิดไม่ได้", str(e))
            return
        job = self._snapshot()
        name = self._default_name()
        self._do(lambda: self._open_worker(job, name))

    def _open_worker(self, job, name):
        out = me.unique_path(os.path.join(me.OUTPUT_DIR, name + ".docx"))
        docx = self._gen_docx(job, out)
        if isinstance(docx, list):
            docx = docx[0]
        wio.open_in_word(docx)
        return ("status", f"เปิดใน Word: {os.path.basename(docx)}")

    def _on_pdf(self):
        if self.busy:
            return
        try:
            self._precheck()
        except Exception as e:
            messagebox.showinfo("ยังส่งออกไม่ได้", str(e))
            return
        job = self._snapshot()
        name = self._default_name()
        out = filedialog.asksaveasfilename(
            title="ส่งออก PDF", defaultextension=".pdf",
            filetypes=[("PDF", "*.pdf")],
            initialdir=me.OUTPUT_DIR, initialfile=name + ".pdf",
        )
        if not out:
            self._set_status("ยกเลิกส่งออก PDF", "#a60")
            return
        self._do(lambda: self._pdf_worker(job, out))

    def _pdf_worker(self, job, out):
        tmp = os.path.join(tempfile.gettempdir(), "_tent_pdf_src.docx")
        docx = self._gen_docx(job, tmp)
        if isinstance(docx, list):
            docx = docx[0]
        pdf = wio.export_pdf(docx, out)
        return ("reveal", os.path.dirname(pdf), f"ส่งออก PDF แล้ว: {pdf}")

    def _on_print(self):
        """กดปุ่ม Print: เช็คข้อมูล -> เลือกเครื่องพิมพ์ -> สั่งพิมพ์ (ใน worker)"""
        if self.busy:
            return
        try:
            self._precheck()
        except Exception as e:
            messagebox.showerror("พิมพ์ไม่ได้", str(e))
            return
        printer = self._ask_printer()
        if printer is None:  # ยกเลิก
            return
        job = self._snapshot()
        self._do(lambda: self._print_worker(job, printer))

    def _precheck(self):
        """ตรวจว่ามีข้อมูลพอจะสร้างป้ายไหม (เรียกบนเธรดหลักก่อนเริ่มงาน)"""
        if self.mode.get() == "single":
            if not self.e_name.get().strip():
                raise ValueError("กรุณากรอกชื่อก่อน")
        elif not self._batch_rows:
            raise ValueError("ยังไม่ได้นำเข้ารายชื่อ (เลือกไฟล์ Excel/CSV ก่อน)")

    def _ask_printer(self):
        """เปิดหน้าต่างเลือกเครื่องพิมพ์ คืนค่าชื่อเครื่องพิมพ์, "" = เครื่องพิมพ์เริ่มต้น, None = ยกเลิก"""
        try:
            printers, default = wio.list_printers()
        except Exception:
            printers, default = [], None
        if not printers:
            if messagebox.askyesno("พิมพ์", "ดึงรายชื่อเครื่องพิมพ์ไม่ได้\nใช้เครื่องพิมพ์เริ่มต้นหรือไม่?"):
                return ""
            return None
        return PrinterDialog(self, printers, default).result

    def _print_worker(self, job, printer=""):
        tmp = os.path.join(tempfile.gettempdir(), "_tent_print_src.docx")
        docx = self._gen_docx(job, tmp)
        if isinstance(docx, list):
            docx = docx[0]
        wio.print_doc(docx, printer or None)
        where = printer if printer else "เครื่องพิมพ์เริ่มต้น"
        return ("status", f"ส่งงานพิมพ์ไปที่: {where}")

    # ---------------------------------------------------------- batch list
    def _choose_list(self):
        path = filedialog.askopenfilename(
            title="เลือกไฟล์รายชื่อ",
            filetypes=[("Excel/CSV", "*.xlsx *.xlsm *.csv"), ("ทุกไฟล์", "*.*")],
        )
        if not path:
            return
        try:
            rows = data_io.load_rows(path)
        except Exception as e:
            messagebox.showerror("อ่านไฟล์ไม่ได้", str(e))
            return
        if not rows:
            messagebox.showwarning("ไม่มีข้อมูล", "ไม่พบรายชื่อในไฟล์ (ตรวจหัวคอลัมน์ ชื่อ/ตำแหน่ง/หน่วยงาน)")
            return
        self._batch_rows = rows
        self._batch_file = path
        self._preview_index = 0
        self._preview_cache.clear()
        self.lbl_batch.configure(text=f"{os.path.basename(path)}  ({len(rows)} คน)")
        self.tree.delete(*self.tree.get_children())
        for r in rows:
            self.tree.insert("", "end", values=(r["name"], r.get("position", ""), r.get("org", "")))
        self._restripe()
        # โหมดนำเข้า -> โชว์แถบเลื่อนดู
        if self.mode.get() == "batch":
            self.navbar.pack(fill="x", pady=(0, 6), before=self.canvas)
        self._update_nav_label()
        self._refresh_buttons()
        self._set_status(f"นำเข้า {len(rows)} คนแล้ว — กด Preview เพื่อดู (เลื่อน ◀ ▶ ดูคนอื่นได้)", "#080")

    # ---- แก้/เพิ่ม/ลบ รายชื่อในตาราง ----
    def _edit_cell(self, event):
        if self.busy:
            return
        rowid = self.tree.identify_row(event.y)
        colid = self.tree.identify_column(event.x)
        if rowid and colid:
            self._open_editor(rowid, colid)

    def _open_editor(self, rowid, colid):
        bbox = self.tree.bbox(rowid, colid)
        if not bbox:
            return
        x, y, w, h = bbox
        col_idx = int(colid[1:]) - 1
        colname = self.tree["columns"][col_idx]
        cur = self.tree.set(rowid, colname)
        entry = ttk.Entry(self.tree, font=UI_FONT)
        entry.place(x=x, y=y, width=w, height=max(h, 22))
        entry.insert(0, cur)
        entry.focus_set()
        entry.select_range(0, "end")

        def commit(_=None):
            if not entry.winfo_exists():
                return
            val = entry.get()
            entry.destroy()
            self.tree.set(rowid, colname, val)
            self._sync_rows_from_tree()

        def cancel(_=None):
            if entry.winfo_exists():
                entry.destroy()

        entry.bind("<Return>", commit)
        entry.bind("<FocusOut>", commit)
        entry.bind("<Escape>", cancel)

    def _add_person(self):
        if self.mode.get() != "batch":
            return
        iid = self.tree.insert("", "end", values=("ชื่อใหม่", "", ""))
        self.tree.selection_set(iid)
        self.tree.see(iid)
        self._sync_rows_from_tree()
        if not self.navbar.winfo_ismapped():
            self.navbar.pack(fill="x", pady=(0, 6), before=self.canvas)
        self.after(60, lambda: self._open_editor(iid, "#1"))  # เริ่มแก้ช่องชื่อทันที

    def _delete_selected(self):
        sel = self.tree.selection()
        if not sel:
            messagebox.showinfo("ลบรายชื่อ", "กรุณาคลิกเลือกแถวที่ต้องการลบก่อน")
            return
        for iid in sel:
            self.tree.delete(iid)
        self._sync_rows_from_tree()
        if not self.tree.get_children():
            self.navbar.pack_forget()
            self.canvas.delete("all")
        self._set_status(f"ลบแล้ว {len(sel)} รายการ", "#080")

    def _clear_all(self):
        """ล้างรายชื่อทั้งหมดในตาราง (ถามยืนยันก่อน เพราะลบทุกแถวพร้อมกัน)"""
        if self.busy:
            return
        n = len(self.tree.get_children())
        if not n:
            self._set_status("ไม่มีรายชื่อให้ล้าง", "#a60")
            return
        if not messagebox.askyesno(
                "ล้างทั้งหมด",
                f"ต้องการล้างรายชื่อทั้งหมด {n} รายการใช่หรือไม่?\n(ลบแล้วเรียกคืนไม่ได้)"):
            return
        self.tree.delete(*self.tree.get_children())
        self._batch_rows = []
        self._batch_file = None
        self._preview_index = 0
        self._preview_cache.clear()
        self.lbl_batch.configure(text="ยังไม่ได้เลือกไฟล์")
        self.navbar.pack_forget()
        self._stop_spinner()
        self.canvas.delete("all")
        self._update_nav_label()
        self._refresh_buttons()
        self._set_status(f"ล้างรายชื่อทั้งหมดแล้ว ({n} รายการ)", "#080")

    def _sync_rows_from_tree(self):
        """อัปเดต _batch_rows ให้ตรงกับตาราง + ล้างแคชพรีวิว"""
        rows = []
        for iid in self.tree.get_children():
            v = self.tree.item(iid, "values")
            rows.append({"name": (v[0] or "").strip(),
                         "position": (v[1] or "").strip() if len(v) > 1 else "",
                         "org": (v[2] or "").strip() if len(v) > 2 else ""})
        self._batch_rows = rows
        self._preview_cache.clear()
        if self._preview_index >= len(rows):
            self._preview_index = max(0, len(rows) - 1)
        self._restripe()
        self._refresh_batch_meta()
        self._update_nav_label()
        self._refresh_buttons()

    def _restripe(self):
        """ทำแถบสลับสีในตาราง (อ่านง่ายขึ้น)"""
        for i, iid in enumerate(self.tree.get_children()):
            self.tree.item(iid, tags=("even" if i % 2 else "odd",))

    def _refresh_batch_meta(self):
        n = len(self.tree.get_children())
        base = os.path.basename(self._batch_file) if self._batch_file else "รายการ"
        self.lbl_batch.configure(text=f"{base}  ({n} คน)")

    # ---------------------------------------------------------- misc
    def _reveal(self, folder):
        try:
            os.startfile(folder)  # type: ignore[attr-defined]
        except Exception:
            pass

    def _on_close(self):
        """ปิด Word instance ที่เปิดค้างไว้ก่อนปิดโปรแกรม"""
        try:
            wio.shutdown_word()
        except Exception:
            pass
        self.destroy()


class PrinterDialog(tk.Toplevel):
    """หน้าต่างเลือกเครื่องพิมพ์ (modal). result = ชื่อเครื่องพิมพ์ หรือ None ถ้ายกเลิก"""

    def __init__(self, parent, printers, default):
        super().__init__(parent)
        self.title("เลือกเครื่องพิมพ์")
        self.result = None
        self.configure(bg=COL_BG)
        self.resizable(False, False)
        self.transient(parent)

        frm = ttk.Frame(self, padding=16)
        frm.pack(fill="both", expand=True)
        ttk.Label(frm, text="เลือกเครื่องพิมพ์ที่ต้องการ:", font=UI_FONT_BOLD).pack(anchor="w")
        self.var = tk.StringVar(value=default or (printers[0] if printers else ""))
        self.combo = ttk.Combobox(frm, textvariable=self.var, values=printers,
                                  state="readonly", width=44, font=UI_FONT)
        self.combo.pack(fill="x", pady=(8, 16))

        btns = ttk.Frame(frm)
        btns.pack(fill="x")
        ttk.Button(btns, text="ยกเลิก", command=self._cancel).pack(side="right")
        ttk.Button(btns, text="🖨 พิมพ์", style="Primary.TButton",
                   command=self._ok).pack(side="right", padx=(0, 8))

        self.bind("<Return>", lambda e: self._ok())
        self.bind("<Escape>", lambda e: self._cancel())
        self.protocol("WM_DELETE_WINDOW", self._cancel)

        self.update_idletasks()
        self._center(parent)
        self.grab_set()
        self.combo.focus_set()
        self.wait_window(self)

    def _ok(self):
        self.result = self.var.get()
        self.destroy()

    def _cancel(self):
        self.result = None
        self.destroy()

    def _center(self, parent):
        try:
            w, h = self.winfo_width(), self.winfo_height()
            x = parent.winfo_rootx() + (parent.winfo_width() - w) // 2
            y = parent.winfo_rooty() + (parent.winfo_height() - h) // 3
            self.geometry(f"+{max(x, 0)}+{max(y, 0)}")
        except Exception:
            pass


def _close_splash():
    """ปิดหน้าโหลด (splash) ของ PyInstaller — มีเฉพาะตอนรันจาก .exe ที่ build พร้อม --splash
    เมื่อรันจากซอร์สจะไม่มีโมดูล pyi_splash จึง no-op อย่างปลอดภัย"""
    try:
        import pyi_splash            # มีเฉพาะใน frozen build ที่ใส่ Splash ไว้
    except Exception:
        return
    try:
        pyi_splash.close()
    except Exception:
        pass


def main():
    app = App()
    app.update()        # บังคับให้หน้าต่างหลักวาดเสร็จก่อน
    _close_splash()     # แล้วค่อยปิดหน้าโหลด -> ผู้ใช้ไม่เห็นจอว่างคั่น
    app.mainloop()


if __name__ == "__main__":
    main()
