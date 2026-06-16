# -*- coding: utf-8 -*-
"""
word_io.py
เชื่อมกับ Microsoft Word ผ่าน COM (pywin32) สำหรับฟีเจอร์ที่ต้องใช้ Word จริง:
- open_in_word     : เปิดไฟล์ใน Word ให้ผู้ใช้ตรวจ/แก้/พิมพ์เอง
- export_pdf       : ส่งออกเป็น PDF
- print_doc        : สั่งพิมพ์ออกเครื่องพิมพ์
- render_preview_png : docx -> PDF (Word) -> PNG หน้าแรก (ผ่าน PyMuPDF) สำหรับพรีวิวในโปรแกรม

ทุกฟังก์ชันปลอดภัยเมื่อเรียกจาก worker thread (CoInitialize/CoUninitialize)
และจะ raise WordError พร้อมข้อความภาษาไทยที่อ่านง่ายเมื่อขาดของหรือมีปัญหา
"""
import os
import tempfile

WD_EXPORT_FORMAT_PDF = 17  # wdExportFormatPDF
WORD_OP_TIMEOUT = 120      # วินาที: กันงาน Word ค้างไม่รู้จบ (เช่น Word เด้ง dialog ค้าง)


class WordError(Exception):
    pass


def _abspath(p):
    return os.path.abspath(p)


# ---------------------------------------------------------------------------
# ตรวจความพร้อมของระบบ (ให้ GUI เปิด/ปิดปุ่มได้)
# ---------------------------------------------------------------------------
def pywin32_available() -> bool:
    try:
        import win32com.client  # noqa: F401
        return True
    except ImportError:
        return False


def fitz_available() -> bool:
    try:
        import fitz  # noqa: F401
        return True
    except ImportError:
        return False


def word_available() -> bool:
    """มี pywin32 และเปิด Word.Application ได้จริงไหม"""
    if not pywin32_available():
        return False
    import win32com.client
    import pythoncom
    pythoncom.CoInitialize()
    word = None
    try:
        word = win32com.client.DispatchEx("Word.Application")
        return True
    except Exception:
        return False
    finally:
        try:
            if word is not None:
                word.Quit()
        except Exception:
            pass
        pythoncom.CoUninitialize()


def _require_pywin32():
    if not pywin32_available():
        raise WordError(
            "ไม่พบ pywin32 (ตัวเชื่อม Word)\nติดตั้งด้วย:  pip install pywin32"
        )


def list_printers():
    """คืนค่า (รายชื่อเครื่องพิมพ์ทั้งหมด, ชื่อเครื่องพิมพ์เริ่มต้น)"""
    _require_pywin32()
    import win32print

    flags = win32print.PRINTER_ENUM_LOCAL | win32print.PRINTER_ENUM_CONNECTIONS
    printers = []
    for p in win32print.EnumPrinters(flags):
        name = p[2]
        if name and name not in printers:
            printers.append(name)
    try:
        default = win32print.GetDefaultPrinter()
    except Exception:
        default = printers[0] if printers else None
    return printers, default


# ---------------------------------------------------------------------------
# Word worker: เปิด Word ค้างไว้ instance เดียวบนเธรดเฉพาะ (เร็วกว่าเปิด/ปิดทุกครั้ง)
# COM ต้องเรียกจากเธรดเดิมที่สร้างมัน จึงรวมงาน Word ทั้งหมดไว้เธรดนี้
# ---------------------------------------------------------------------------
import threading
import queue


class _WordWorker:
    def __init__(self):
        self._q = queue.Queue()
        self._thread = None
        self._app = None
        self._lock = threading.Lock()

    def _ensure_thread(self):
        if self._thread is None or not self._thread.is_alive():
            self._thread = threading.Thread(target=self._run, name="WordWorker", daemon=True)
            self._thread.start()

    def submit(self, fn, timeout=None):
        """ส่งงาน fn(worker) ไปรันบนเธรด Word แล้วรอผลลัพธ์ (โยน exception กลับมาถ้ามี)
        ถ้าเกิน timeout วินาทีแล้วยังไม่ตอบ -> โยน WordError แทนการรอค้างตลอดไป"""
        with self._lock:
            self._ensure_thread()
            box = queue.Queue()
            self._q.put((fn, box))
        try:
            ok, val = box.get(timeout=timeout)
        except queue.Empty:
            raise WordError(
                "Word ไม่ตอบสนอง (หมดเวลารอ)\n"
                "โปรดปิดหน้าต่าง Word ที่อาจค้างอยู่ แล้วลองใหม่อีกครั้ง"
            )
        if ok:
            return val
        raise val

    def app(self):
        """คืน Word.Application ที่เปิดค้างไว้ (สร้างครั้งแรกครั้งเดียว) — เรียกบนเธรด Word เท่านั้น"""
        import win32com.client
        if self._app is None:
            self._app = win32com.client.DispatchEx("Word.Application")
            self._app.Visible = False
            try:
                self._app.DisplayAlerts = 0
            except Exception:
                pass
        return self._app

    def _run(self):
        import pythoncom
        pythoncom.CoInitialize()
        try:
            while True:
                fn, box = self._q.get()
                if fn is None:
                    break
                try:
                    box.put((True, fn(self)))
                except Exception as e:
                    # Word อาจค้าง/ถูกปิด -> รีเซ็ตให้สร้างใหม่รอบหน้า
                    self._app = None
                    box.put((False, e))
        finally:
            if self._app is not None:
                try:
                    self._app.Quit()
                except Exception:
                    pass
                self._app = None
            pythoncom.CoUninitialize()

    def shutdown(self):
        if self._thread is not None and self._thread.is_alive():
            self._q.put((None, None))
            self._thread.join(timeout=5)
        self._thread = None


_worker = _WordWorker()


def shutdown_word():
    """ปิด Word instance ที่เปิดค้างไว้ (เรียกตอนปิดโปรแกรม)"""
    try:
        _worker.shutdown()
    except Exception:
        pass


# ---------------------------------------------------------------------------
# ฟีเจอร์ Word
# ---------------------------------------------------------------------------
def open_in_word(path):
    """เปิดไฟล์ใน Word แบบมองเห็น แล้วปล่อยให้ผู้ใช้จัดการเอง"""
    _require_pywin32()
    import win32com.client
    import pythoncom

    path = _abspath(path)
    if not os.path.exists(path):
        raise WordError(f"ไม่พบไฟล์: {path}")

    pythoncom.CoInitialize()
    try:
        word = win32com.client.Dispatch("Word.Application")  # ใช้ instance เดิมถ้ามี
        word.Visible = True
        word.Documents.Open(path)
        try:
            word.Activate()
        except Exception:
            pass
    except WordError:
        raise
    except Exception as e:
        raise WordError(f"เปิดใน Word ไม่สำเร็จ: {e}")
    finally:
        pythoncom.CoUninitialize()


def export_pdf(docx_path, pdf_path=None):
    """ส่งออก .docx เป็น .pdf ด้วย Word instance ที่เปิดค้างไว้. คืนค่า path ของ PDF"""
    _require_pywin32()
    docx_path = _abspath(docx_path)
    if not os.path.exists(docx_path):
        raise WordError(f"ไม่พบไฟล์: {docx_path}")
    if pdf_path is None:
        pdf_path = os.path.splitext(docx_path)[0] + ".pdf"
    pdf_path = _abspath(pdf_path)

    def task(w):
        app = w.app()
        doc = app.Documents.Open(docx_path, ReadOnly=True)
        try:
            doc.ExportAsFixedFormat(pdf_path, WD_EXPORT_FORMAT_PDF)
        finally:
            doc.Close(False)
        return pdf_path

    try:
        return _worker.submit(task, timeout=WORD_OP_TIMEOUT)
    except WordError:
        raise
    except Exception as e:
        raise WordError(f"ส่งออก PDF ไม่สำเร็จ: {e}")


def print_doc(path, printer=None):
    """สั่งพิมพ์ไฟล์ออกเครื่องพิมพ์ (default printer หรือชื่อที่ระบุ) ด้วย Word ที่เปิดค้างไว้"""
    _require_pywin32()
    path = _abspath(path)
    if not os.path.exists(path):
        raise WordError(f"ไม่พบไฟล์: {path}")

    def task(w):
        app = w.app()
        if printer:
            try:
                app.ActivePrinter = printer
            except Exception:
                pass
        doc = app.Documents.Open(path, ReadOnly=True)
        try:
            doc.PrintOut(Background=False)  # รอ spool เสร็จก่อนปิด
        finally:
            doc.Close(False)
        return True

    try:
        _worker.submit(task, timeout=WORD_OP_TIMEOUT)
    except WordError:
        raise
    except Exception as e:
        raise WordError(f"สั่งพิมพ์ไม่สำเร็จ: {e}")


def render_preview_png(docx_path, png_path=None, dpi=110):
    """docx -> PDF (Word) -> PNG หน้าแรก (PyMuPDF). คืนค่า path ของ PNG"""
    if not fitz_available():
        raise WordError(
            "ไม่พบ PyMuPDF (ตัวเรนเดอร์รูป)\nติดตั้งด้วย:  pip install pymupdf"
        )
    import fitz

    docx_path = _abspath(docx_path)
    tmp_pdf = os.path.join(tempfile.gettempdir(), "_tent_preview.pdf")
    export_pdf(docx_path, tmp_pdf)

    if png_path is None:
        png_path = os.path.join(tempfile.gettempdir(), "_tent_preview.png")
    png_path = _abspath(png_path)

    pdf = fitz.open(tmp_pdf)
    try:
        page = pdf[0]
        zoom = dpi / 72.0
        pix = page.get_pixmap(matrix=fitz.Matrix(zoom, zoom))
        pix.save(png_path)
    finally:
        pdf.close()
    return png_path


if __name__ == "__main__":
    print("pywin32:", pywin32_available())
    print("fitz   :", fitz_available())
    print("word   :", word_available())
