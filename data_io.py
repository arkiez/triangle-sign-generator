# -*- coding: utf-8 -*-
"""
data_io.py
อ่านรายชื่อจากไฟล์ Excel (.xlsx) หรือ CSV สำหรับโหมด batch
คืนค่าเป็น list ของ dict {name, position, org}
รองรับหัวคอลัมน์ทั้งไทยและอังกฤษ ถ้าไม่พบหัวคอลัมน์จะถือตามลำดับ: ชื่อ, ตำแหน่ง, หน่วยงาน
"""
import csv
import os

NAME_KEYS = {"name", "ชื่อ", "ชื่อ-สกุล", "ชื่อ-นามสกุล", "ชื่อสกุล", "fullname", "full name"}
POS_KEYS = {"position", "pos", "ตำแหน่ง"}
ORG_KEYS = {"org", "organization", "หน่วยงาน", "สังกัด", "department"}


def _norm(s):
    return (str(s) if s is not None else "").strip().lower()


def _build_colmap(header):
    """คืนค่า dict {field: index} จากแถวหัวตาราง; ค่า None ถ้าหาไม่เจอ"""
    cmap = {"name": None, "position": None, "org": None}
    for i, cell in enumerate(header):
        key = _norm(cell)
        if key in NAME_KEYS and cmap["name"] is None:
            cmap["name"] = i
        elif key in POS_KEYS and cmap["position"] is None:
            cmap["position"] = i
        elif key in ORG_KEYS and cmap["org"] is None:
            cmap["org"] = i
    return cmap


def _rows_from_matrix(matrix):
    """matrix: list ของ list (รวมหัวตาราง). คืนค่า list ของ dict"""
    matrix = [r for r in matrix if any((str(c).strip() if c is not None else "") for c in r)]
    if not matrix:
        return []

    header = matrix[0]
    cmap = _build_colmap(header)

    if cmap["name"] is None:
        # ไม่พบหัวคอลัมน์ -> ถือตามลำดับคอลัมน์ และนับแถวแรกเป็นข้อมูลด้วย
        cmap = {"name": 0, "position": 1, "org": 2}
        data_rows = matrix
    else:
        data_rows = matrix[1:]

    def get(row, idx):
        if idx is None or idx >= len(row):
            return ""
        v = row[idx]
        return str(v).strip() if v is not None else ""

    out = []
    for row in data_rows:
        name = get(row, cmap["name"])
        if not name:
            continue
        out.append(
            {
                "name": name,
                "position": get(row, cmap["position"]),
                "org": get(row, cmap["org"]),
            }
        )
    return out


def load_rows(path):
    """อ่านไฟล์ .xlsx หรือ .csv -> list ของ dict {name, position, org}"""
    ext = os.path.splitext(path)[1].lower()
    if ext in (".xlsx", ".xlsm"):
        return _load_xlsx(path)
    if ext in (".csv", ".txt"):
        return _load_csv(path)
    raise ValueError(f"ไม่รองรับไฟล์ชนิด {ext} (รองรับ .xlsx และ .csv)")


def _load_xlsx(path):
    try:
        import openpyxl
    except ImportError:
        raise RuntimeError("ต้องติดตั้ง openpyxl ก่อน:  pip install openpyxl")
    wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    ws = wb.active
    matrix = [list(r) for r in ws.iter_rows(values_only=True)]
    wb.close()
    return _rows_from_matrix(matrix)


def _load_csv(path):
    # utf-8-sig รองรับ BOM ที่ Excel ใส่มาเวลา Save As CSV
    for enc in ("utf-8-sig", "utf-8", "cp874"):
        try:
            with open(path, newline="", encoding=enc) as f:
                matrix = list(csv.reader(f))
            return _rows_from_matrix(matrix)
        except UnicodeDecodeError:
            continue
    raise RuntimeError("อ่านไฟล์ CSV ไม่ได้ (ลองบันทึกเป็น UTF-8)")
