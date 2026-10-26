"""ST Print Backoffice & ERP Web Dashboard Server.

Modules:
1. รายรับ (โอนเป็นหลัก): ซิงก์ยอดโอนมัดจำ + ยอดโอนจ่ายเต็ม/ส่วนต่าง จาก Google Sheet อัตโนมัติ + บันทึกรายรับเพิ่มเติม
2. รายจ่าย / เจ้าหนี้ / ค่าใช้จ่าย: บันทึกรายจ่าย, เจ้าหนี้ค้างจ่าย, วันครบกำหนดชำระ, แยกหมวดหมู่, สรุปกำไรสุทธิ
3. สต๊อก (สติ๊กเกอร์ / จำนวน / ประเภท / บรรจุภัณฑ์): ตัดยอดแผ่นสติ๊กเกอร์ที่ใช้จริงจากชีตออเดอร์อัตโนมัติ + รับเข้า/เบิกออก + แจ้งเตือนของใกล้หมด
4. ลูกค้า (CRM & Order Tracker): ประวัติการสั่งซื้อรายคน, รายละเอียดการสั่ง, ยอดมัดจำ/จ่ายเต็ม/ส่วนต่าง, สถานะการเตรียมผลิต และสถานะการจัดส่ง + เลขพัสดุ
"""
from __future__ import annotations

import base64
import csv
import datetime as dt
import io
import json
import os
import re
import shutil
import socket
import struct
import subprocess
import sys
import threading
import time
import urllib.parse
import urllib.request
import uuid
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path
from typing import Any

BASE_DIR = Path(__file__).resolve().parent
DATA_DIR = BASE_DIR / "data"
DATA_DIR.mkdir(parents=True, exist_ok=True)

STORE_PATH = DATA_DIR / "erp_store.json"
SHEET_CACHE_PATH = DATA_DIR / "sheet_orders_cache.json"
ADMIN_CACHE_PATH = DATA_DIR / "admin_page_cache.json"
MYORDER_CACHE_PATH = DATA_DIR / "myorder_shipping_cache.json"
MYORDER_TOKEN_PATH = DATA_DIR / "myorder_token.txt"
XLSX_CACHE_PATH = DATA_DIR / "latest_sheet.xlsx"

MYORDER_TEAM_ID = "62e0f2acb2da9100180e480f"
MYORDER_OM_BASE = "https://mod-om-gateway.myorder.ai"
MYORDER_CDP_PORT = 9225
MYORDER_USER_AGENT = (
    "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
    "AppleWebKit/537.36 (KHTML, like Gecko) "
    "Chrome/131.0.0.0 Safari/537.36"
)

GSHEET_ID = "1cQzf9GhR5MdW8gDw3dTopLFAP8Z0gUiLeeqcy6EcM0c"
GSHEET_OCT_GID = "2140811773"
GSHEET_ADMIN_ID = "1xIo-PlD4j1aPrpK6YRGW50cdWWJpMZGnEO0LCKcjwVY"
GSHEET_ADMIN_OCT_GID = "1846638698"
PORT = int(os.environ.get("STPRINT_ERP_PORT", "3580"))

ADMIN_PAGES = [
    "STprint",
    "PLP",
    "โรงงานสติ๊กเกอร์25",
    "ถูกดีจริง",
    "ถูกเว้ย25",
    "ฉลาก25",
    "พิมพ์สวย",
    "ปริ้นฉลาก",
]
ADMIN_NAMES = ["จูน", "เอม", "ปอนด์", "มาย"]

LAYOUT_NEW = {
    "seq": 1, "entry": 2, "dep_date": 3, "dep_amt": 4, "dep_time": 5,
    "full_date": 6, "full_amt": 7, "full_time": 8, "bank": None, "paid": 9,
    "note": 10, "new_cust": 11, "cust_file": 12, "old_file": 13, "sales": 14,
    "channel": 15, "fb": 16, "job": 17, "designer": 18, "print_time": 19,
    "color": 20, "size": 21, "sheets": 22, "bonus": 23, "qty": 24,
    "material": 25, "phone": 26, "recipient": 27, "address": 28, "st": 29,
}
LAYOUT_OLD = {
    "seq": 1, "entry": None, "dep_date": 2, "dep_amt": 3, "dep_time": 4,
    "full_date": 5, "full_amt": 6, "full_time": 7, "bank": 8, "paid": 9,
    "note": 10, "new_cust": 11, "cust_file": 12, "old_file": 13, "sales": 14,
    "channel": 15, "fb": 16, "job": 17, "designer": 18, "print_time": 19,
    "color": 20, "size": 21, "sheets": 22, "bonus": 23, "qty": 24,
    "material": 25, "phone": 26, "recipient": 27, "address": 28, "st": 29,
}

DEFAULT_STOCK_ITEMS = [
    {
        "id": "stk-pp-gloss",
        "code": "STK-001",
        "name": "สติ๊กเกอร์ PP ขาวเงา (A3+)",
        "category": "สติ๊กเกอร์",
        "material_match": ["ขาวเงา"],
        "unit": "แผ่น",
        "initial_qty": 50000,
        "min_qty": 3000,
        "cost_per_unit": 4.5,
        "note": "เนื้อหลักของร้าน (กันน้ำ 100%)",
    },
    {
        "id": "stk-pp-clear",
        "code": "STK-002",
        "name": "สติ๊กเกอร์ PP ใส (A3+)",
        "category": "สติ๊กเกอร์",
        "material_match": ["ใส"],
        "unit": "แผ่น",
        "initial_qty": 5000,
        "min_qty": 500,
        "cost_per_unit": 5.0,
        "note": "สติ๊กเกอร์ใสกันน้ำ",
    },
    {
        "id": "stk-pp-matte",
        "code": "STK-003",
        "name": "สติ๊กเกอร์ PP ขาวด้าน (A3+)",
        "category": "สติ๊กเกอร์",
        "material_match": ["ขาวด้าน"],
        "unit": "แผ่น",
        "initial_qty": 3000,
        "min_qty": 400,
        "cost_per_unit": 4.8,
        "note": "สติ๊กเกอร์ขาวด้านกันน้ำ",
    },
    {
        "id": "stk-pvc-gloss",
        "code": "STK-004",
        "name": "สติ๊กเกอร์ PVC ขาวเงา (A3+)",
        "category": "สติ๊กเกอร์",
        "material_match": ["PVC ขาวเงา"],
        "unit": "แผ่น",
        "initial_qty": 2500,
        "min_qty": 300,
        "cost_per_unit": 6.5,
        "note": "เนื้อ PVC ยืดหยุ่นสูง",
    },
    {
        "id": "stk-pvc-matte",
        "code": "STK-005",
        "name": "สติ๊กเกอร์ PVC ขาวด้าน (A3+)",
        "category": "สติ๊กเกอร์",
        "material_match": ["PVC ขาวด้าน"],
        "unit": "แผ่น",
        "initial_qty": 1500,
        "min_qty": 200,
        "cost_per_unit": 6.8,
        "note": "เนื้อ PVC ขาวด้าน",
    },
    {
        "id": "stk-pvc-clear",
        "code": "STK-006",
        "name": "สติ๊กเกอร์ PVC ใส (A3+)",
        "category": "สติ๊กเกอร์",
        "material_match": ["PVC ใส"],
        "unit": "แผ่น",
        "initial_qty": 1500,
        "min_qty": 200,
        "cost_per_unit": 6.8,
        "note": "เนื้อ PVC ใส",
    },
    {
        "id": "stk-art-paper",
        "code": "STK-007",
        "name": "สติ๊กเกอร์กระดาษอาร์ต (A3+)",
        "category": "สติ๊กเกอร์",
        "material_match": ["กระดาษอาร์ต", "กระดาษ"],
        "unit": "แผ่น",
        "initial_qty": 2000,
        "min_qty": 300,
        "cost_per_unit": 2.8,
        "note": "สติ๊กเกอร์กระดาษ",
    },
    {
        "id": "stk-vinyl",
        "code": "STK-008",
        "name": "สติ๊กเกอร์ไวนิล",
        "category": "สติ๊กเกอร์",
        "material_match": ["ไวนิล"],
        "unit": "แผ่น",
        "initial_qty": 500,
        "min_qty": 100,
        "cost_per_unit": 12.0,
        "note": "งานป้าย/ไวนิล",
    },
    {
        "id": "pkg-bag-a3",
        "code": "PKG-001",
        "name": "ซองพลาสติกใสใส่สติ๊กเกอร์ A3+",
        "category": "บรรจุภัณฑ์",
        "material_match": ["__ORDER_COUNT__"],
        "unit": "ใบ",
        "initial_qty": 5000,
        "min_qty": 500,
        "cost_per_unit": 1.5,
        "note": "ตัดอัตโนมัติ 1 ใบต่อ 1 ออเดอร์จัดส่ง",
    },
    {
        "id": "pkg-board-a3",
        "code": "PKG-002",
        "name": "แผ่นกระดาษแข็งดามหลังกันยับ A3+",
        "category": "บรรจุภัณฑ์",
        "material_match": ["__ORDER_COUNT__"],
        "unit": "แผ่น",
        "initial_qty": 5000,
        "min_qty": 500,
        "cost_per_unit": 3.0,
        "note": "ตัดอัตโนมัติ 1 แผ่นต่อ 1 ออเดอร์จัดส่ง",
    },
    {
        "id": "pkg-outer-env",
        "code": "PKG-003",
        "name": "ซองไปรษณีย์พลาสติกกันน้ำ / กล่องพัสดุ",
        "category": "บรรจุภัณฑ์",
        "material_match": ["__ORDER_COUNT__"],
        "unit": "ชิ้น",
        "initial_qty": 5000,
        "min_qty": 500,
        "cost_per_unit": 2.5,
        "note": "ตัดอัตโนมัติ 1 ชิ้นต่อ 1 ออเดอร์จัดส่ง",
    },
    {
        "id": "pkg-tape",
        "code": "PKG-004",
        "name": "เทปกาวแพ็กพัสดุ",
        "category": "บรรจุภัณฑ์",
        "material_match": [],
        "unit": "ม้วน",
        "initial_qty": 120,
        "min_qty": 20,
        "cost_per_unit": 25.0,
        "note": "เบิกใช้ตามม้วนจริง",
    },
]

DEFAULT_USERS = [
    {
        "id": "usr-owner",
        "username": "admin",
        "password": "1234",
        "display_name": "👑 เจ้าของร้าน (ดู/แก้ไขได้ทุกหน้า)",
        "tabs": ["overview", "income", "expense", "stock", "customer", "adminboard"],
        "can_edit": True,
        "hide_money": False,
        "can_manage_users": True,
    },
    {
        "id": "usr-stock",
        "username": "stock",
        "password": "1234",
        "display_name": "📦 ฝ่ายผลิต / แพ็กส่ง / เช็คสต๊อก",
        "tabs": ["stock", "customer"],
        "can_edit": True,
        "hide_money": True,
        "can_manage_users": False,
    },
    {
        "id": "usr-viewer",
        "username": "viewer",
        "password": "1234",
        "display_name": "👀 พนักงาน (ดูอย่างเดียว - View Only)",
        "tabs": ["stock", "customer"],
        "can_edit": False,
        "hide_money": True,
        "can_manage_users": False,
    },
]

_LOCK = threading.Lock()


def _load_store() -> dict[str, Any]:
    default_store = {
        "manual_incomes": [],
        "expenses": [],
        "stock_items": DEFAULT_STOCK_ITEMS,
        "stock_movements": [],
        "order_overrides": {},
        "users": DEFAULT_USERS,
        "last_sync_ts": 0.0,
    }
    if not STORE_PATH.exists():
        _save_store(default_store)
        return default_store
    try:
        data = json.loads(STORE_PATH.read_text(encoding="utf-8"))
        for k, v in default_store.items():
            if k not in data:
                data[k] = v
        if not data.get("stock_items"):
            data["stock_items"] = DEFAULT_STOCK_ITEMS
        if not data.get("users"):
            data["users"] = DEFAULT_USERS
        return data
    except Exception:
        return default_store


def _save_store(data: dict[str, Any]) -> None:
    STORE_PATH.write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8")
    try:
        _publish_cloud_state_async()
    except Exception:
        pass


def _norm_date(v: Any, fallback_month: int = 10, fallback_year: int = 2026) -> str | None:
    if v is None:
        return None
    if isinstance(v, dt.datetime):
        d = v.date()
    elif isinstance(v, dt.date):
        d = v
    else:
        s = str(v).strip()
        if not s or s in ("**", "-", "None"):
            return None
        m = re.match(r"^(\d{4})-(\d{2})-(\d{2})", s)
        if m:
            d = dt.date(int(m.group(1)), int(m.group(2)), int(m.group(3)))
        else:
            m2 = re.match(r"^(\d{1,2})[/-](\d{1,2})[/-](\d{2,4})", s)
            if not m2:
                return None
            day, mon, yr = int(m2.group(1)), int(m2.group(2)), int(m2.group(3))
            if yr > 2400:
                yr -= 543
            elif yr < 100:
                yr = 2026 if yr == 69 else (2000 + yr)
            try:
                d = dt.date(yr, mon, day)
            except Exception:
                return None
    # Fix Buddhist short year 69 parsed by Excel as 1969 -> 2026
    if 1965 <= d.year <= 1975:
        try:
            d = d.replace(year=d.year + 57)
        except ValueError:
            d = dt.date(d.year + 57, d.month, 28)
    elif d.year > 2400:
        try:
            d = d.replace(year=d.year - 543)
        except ValueError:
            pass
    return d.isoformat()


def _norm_time(v: Any) -> str:
    if v is None:
        return ""
    if isinstance(v, dt.datetime):
        return v.strftime("%H:%M")
    if isinstance(v, dt.time):
        return v.strftime("%H:%M")
    s = str(v).strip()
    if not s or s in ("**", "-", "None"):
        return ""
    m = re.match(r"^(\d{1,2})[:.](\d{2})", s)
    if m:
        return f"{int(m.group(1)):02d}:{m.group(2)}"
    return s


def _parse_float(v: Any) -> float:
    if v is None or isinstance(v, bool):
        return 0.0
    if isinstance(v, (int, float)):
        return float(v)
    s = str(v).strip().replace(",", "")
    if not s or s in ("**", "-", "None"):
        return 0.0
    try:
        return float(s)
    except Exception:
        return 0.0


def _classify_prep_status(designer: str, raw_print_time: Any) -> tuple[str, str]:
    """Return (status_code, status_label) from sheet columns R (คนทำไฟล์) and S (เวลาส่งปริ้น)."""
    des = str(designer or "").strip()
    if des in ("None", "-", "**"):
        des = ""
    if isinstance(raw_print_time, (dt.datetime, dt.date, dt.time)):
        if isinstance(raw_print_time, (dt.datetime, dt.date)):
            d_iso = _norm_date(raw_print_time) or ""
            return "printed", f"ส่งปริ้นแล้ว ({d_iso})" if d_iso else "ส่งปริ้นแล้ว"
        return "printed", f"ส่งปริ้นแล้ว ({raw_print_time.strftime('%H:%M')})"
    pt = str(raw_print_time or "").strip()
    if pt in ("None", "-", "**"):
        pt = ""
    if pt:
        if re.match(r"^\d{4}-\d{2}-\d{2}", pt):
            d_iso = _norm_date(pt) or pt[:10]
            return "printed", f"ส่งปริ้นแล้ว ({d_iso})"
        if re.match(r"^\d{1,2}[:.]\d{2}", pt):
            return "printed", f"ส่งปริ้นแล้ว ({pt})"
        if "ผ่าน" in pt:
            return "approved", f"แบบผ่านแล้ว ({pt})"
        if "รอตรวจ" in pt:
            return "review", "รอลูกค้าตรวจแบบ"
        if "แก้" in pt:
            return "revision", f"รอแก้แบบ ({pt})"
        return "in_progress", pt
    if des:
        return "designing", f"กำลังทำไฟล์ ({des})"
    return "waiting", "รอทำไฟล์"


def _norm_customer_key(fb: str, phone: str, recipient: str) -> str:
    ph = re.sub(r"\D+", "", str(phone or ""))
    if len(ph) >= 9:
        return f"tel:{ph[:10]}"
    f = re.sub(r"\s+", "", str(fb or "")).lower()
    if f and f not in ("-", "**", "none"):
        return f"fb:{f}"
    r = re.sub(r"\s+", "", str(recipient or "")).lower()
    return f"rec:{r}" if r else "unknown"


def sync_from_google_sheet(all_months: bool = False) -> dict[str, Any]:
    """Download and parse orders & transfers from Google Sheet."""
    import openpyxl

    if all_months:
        url = f"https://docs.google.com/spreadsheets/d/{GSHEET_ID}/export?format=xlsx"
        timeout = 45
    else:
        url = f"https://docs.google.com/spreadsheets/d/{GSHEET_ID}/export?format=xlsx&gid={GSHEET_OCT_GID}"
        timeout = 25

    req = urllib.request.Request(url, headers={"User-Agent": "Mozilla/5.0"})
    with urllib.request.urlopen(req, timeout=timeout) as resp:
        raw_bytes = resp.read()
    XLSX_CACHE_PATH.write_bytes(raw_bytes)

    wb = openpyxl.load_workbook(io.BytesIO(raw_bytes), data_only=True, read_only=True)
    orders: list[dict[str, Any]] = []
    incomes: list[dict[str, Any]] = []

    _THAI_MONTH_MAP = {
        "ม.ค": "01", "ก.พ": "02", "มี.ค": "03", "เม.ย": "04",
        "พ.ค": "05", "มิ.ย": "06", "ก.ค": "07", "ส.ค": "08",
        "ก.ย": "09", "ต.ค": "10", "พ.ย": "11", "ธ.ค": "12",
    }

    def _infer_sheet_month(s_name: str) -> str:
        for th_m, mm in _THAI_MONTH_MAP.items():
            if th_m in s_name:
                return f"2026-{mm}"
        return ""

    try:
        for sheet_name in wb.sheetnames:
            ws = wb[sheet_name]
            sheet_month = _infer_sheet_month(sheet_name)
            rows_iter = ws.iter_rows(values_only=True)
            try:
                hdr_row = next(rows_iter)
            except StopIteration:
                continue
            header_b = str((hdr_row[1] if len(hdr_row) > 1 else "") or "").strip()
            lay = LAYOUT_OLD if header_b == "วันที่" else LAYOUT_NEW

            last_primary_order: dict[str, Any] | None = None
            last_valid_date: str | None = None

            for r_idx, row_vals in enumerate(rows_iter, start=2):
                def g(key: str) -> Any:
                    c = lay.get(key)
                    if not c or c - 1 >= len(row_vals):
                        return None
                    return row_vals[c - 1]

                fb = str(g("fb") or "").strip()
                job = str(g("job") or "").strip()
                if fb in ("None", "**"):
                    fb = ""
                if job in ("None", "**"):
                    job = ""

                dep_amt = _parse_float(g("dep_amt"))
                full_amt = _parse_float(g("full_amt"))
                raw_dep_date = _norm_date(g("dep_date"))
                raw_full_date = _norm_date(g("full_date"))
                print_date = _norm_date(g("print_time"))

                dep_date = raw_dep_date
                full_date = raw_full_date
                if sheet_month:
                    q_fallback = (print_date if (print_date and print_date.startswith(sheet_month)) else f"{sheet_month}-01")
                    if dep_date and not dep_date.startswith(sheet_month):
                        dep_date = q_fallback
                    if full_date and not full_date.startswith(sheet_month):
                        full_date = q_fallback

                if full_date:
                    last_valid_date = full_date
                elif dep_date:
                    last_valid_date = dep_date

                if not fb and not job and dep_amt <= 0 and full_amt <= 0:
                    continue

                raw_seq = str(g("seq") or "").strip()
                is_sub_row = raw_seq == "**" or str(g("full_amt") or "").strip() == "**"

                dep_time = _norm_time(g("dep_time"))
                full_time = _norm_time(g("full_time"))
                paid_flag = str(g("paid") or "").strip().lower() in ("true", "1", "yes", "✓", "✔")
                note = str(g("note") or "").strip()
                if note == "None":
                    note = ""

                is_new_cust = str(g("new_cust") or "").strip().lower() in ("ai", "true", "1", "yes", "✓", "✔")
                is_cust_file = str(g("cust_file") or "").strip().lower() in ("true", "1", "yes", "✓", "✔", "ai")
                is_old_file = str(g("old_file") or "").strip().lower() in ("true", "1", "yes", "✓", "✔", "ai")
                if is_new_cust:
                    order_type = "ลูกค้าใหม่ (ออกแบบ)"
                    order_type_code = "new_cust"
                elif is_cust_file:
                    order_type = "ไฟล์ลูกค้า"
                    order_type_code = "cust_file"
                elif is_old_file:
                    order_type = "ไฟล์เก่า"
                    order_type_code = "old_file"
                else:
                    order_type = "ทั่วไป"
                    order_type_code = "other"

                sales = str(g("sales") or "").strip()
                if sales in ("None", "**", "-"):
                    sales = ""
                channel = str(g("channel") or "").strip()
                if channel in ("None", "**", "-"):
                    channel = ""
                designer = str(g("designer") or "").strip()
                if designer in ("None", "**", "-"):
                    designer = ""
                prep_code, prep_label = _classify_prep_status(designer, g("print_time"))

                size_str = str(g("size") or "").strip()
                if size_str in ("None", "**", "-"):
                    size_str = ""
                sheets_num = _parse_float(g("sheets"))
                bonus_num = _parse_float(g("bonus"))
                qty_text = str(g("qty") or "").strip()
                if qty_text in ("None", "**", "-"):
                    qty_text = f"{int(sheets_num) if sheets_num.is_integer() else sheets_num} แผ่น" if sheets_num > 0 else ""

                material = str(g("material") or "").strip()
                if material in ("None", "**", "-", ""):
                    material = "ขาวเงา" if (sheets_num > 0 or bonus_num > 0 or job) else ""

                phone = str(g("phone") or "").strip()
                if phone in ("None", "**", "-"):
                    phone = ""
                recipient = str(g("recipient") or "").strip()
                if recipient in ("None", "**", "-"):
                    recipient = ""
                address = str(g("address") or "").strip()
                if address in ("None", "**", "-"):
                    address = ""
                st_code = str(g("st") or "").strip()
                if st_code in ("None", "**", "-"):
                    st_code = ""

                order_date = full_date or dep_date or last_valid_date or dt.date.today().isoformat()
                order_id = f"{sheet_name}:R{r_idx}"

                # Inherit customer phone/address/recipient from primary row if this is a '**' continuation row
                if is_sub_row and last_primary_order and (not fb or fb == last_primary_order["fb"]):
                    if not fb:
                        fb = last_primary_order["fb"]
                    if not phone:
                        phone = last_primary_order["phone"]
                    if not recipient:
                        recipient = last_primary_order["recipient"]
                    if not address:
                        address = last_primary_order["address"]
                    if not channel:
                        channel = last_primary_order["channel"]
                    if not sales:
                        sales = last_primary_order["sales"]

                order_rec = {
                    "id": order_id,
                    "sheet_name": sheet_name,
                    "row": r_idx,
                    "seq": raw_seq if raw_seq not in ("None", "**") else "",
                    "is_sub_row": is_sub_row,
                    "date": order_date,
                    "dep_date": dep_date or "",
                    "dep_amt": dep_amt,
                    "dep_time": dep_time,
                    "full_date": full_date or "",
                    "full_amt": full_amt,
                    "full_time": full_time,
                    "total_amt": dep_amt + full_amt,
                    "paid": paid_flag,
                    "note": note,
                    "order_type": order_type,
                    "order_type_code": order_type_code,
                    "sales": sales,
                    "channel": channel,
                    "fb": fb,
                    "job": job,
                    "designer": designer,
                    "prep_status_code": prep_code,
                    "prep_status_label": prep_label,
                    "size": size_str,
                    "sheets": sheets_num,
                    "bonus": bonus_num,
                    "total_sheets": sheets_num + bonus_num,
                    "qty_text": qty_text,
                    "material": material,
                    "phone": phone,
                    "recipient": recipient,
                    "address": address,
                    "st_code": st_code,
                    "customer_key": _norm_customer_key(fb, phone, recipient),
                }
                orders.append(order_rec)
                if not is_sub_row and fb:
                    last_primary_order = order_rec

                # Generate transfer income records ("โอนเป็นหลัก")
                if dep_amt > 0:
                    incomes.append({
                        "id": f"inc-dep-{order_id}",
                        "source": "sheet",
                        "transfer_type": "deposit",
                        "transfer_type_label": "โอนมัดจำ",
                        "date": dep_date or order_date,
                        "time": dep_time,
                        "amount": dep_amt,
                        "method": "โอนเงิน",
                        "verified": paid_flag,
                        "fb": fb,
                        "recipient": recipient,
                        "job": job,
                        "channel": channel,
                        "sales": sales,
                        "order_id": order_id,
                        "row": r_idx,
                        "sheet_name": sheet_name,
                        "note": note,
                    })
                if full_amt > 0:
                    incomes.append({
                        "id": f"inc-full-{order_id}",
                        "source": "sheet",
                        "transfer_type": "full_or_balance",
                        "transfer_type_label": "โอนส่วนต่าง (มีมัดจำ)" if dep_amt > 0 else "โอนยอดเต็ม",
                        "date": full_date or order_date,
                        "time": full_time,
                        "amount": full_amt,
                        "method": "โอนเงิน",
                        "verified": paid_flag,
                        "fb": fb,
                        "recipient": recipient,
                        "job": job,
                        "channel": channel,
                        "sales": sales,
                        "order_id": order_id,
                        "row": r_idx,
                        "sheet_name": sheet_name,
                        "note": note,
                    })
    finally:
        wb.close()

    cache_payload = {
        "synced_at": dt.datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        "orders": orders,
        "incomes": incomes,
    }
    SHEET_CACHE_PATH.write_text(json.dumps(cache_payload, ensure_ascii=False), encoding="utf-8")
    return cache_payload


def sync_admin_page_sheet() -> dict[str, Any]:
    url = f"https://docs.google.com/spreadsheets/d/{GSHEET_ADMIN_ID}/export?format=csv&gid={GSHEET_ADMIN_OCT_GID}"
    req = urllib.request.Request(url, headers={"User-Agent": "Mozilla/5.0"})
    with urllib.request.urlopen(req, timeout=30) as resp:
        raw_csv = resp.read().decode("utf-8", errors="replace")
    rows = list(csv.reader(io.StringIO(raw_csv)))

    def cell(r: list[str], idx: int) -> str:
        return r[idx].strip() if idx < len(r) else ""

    days_raw: dict[str, dict[str, Any]] = {}
    admin_offsets = [("จูน", 1), ("เอม", 42), ("ปอนด์", 83)]

    for r_idx in range(3, min(34, len(rows))):
        r = rows[r_idx]
        d_str = _norm_date(cell(r, 0), 10)
        if not d_str:
            continue
        day_admins: dict[str, dict[str, dict[str, float]]] = {}
        for adm, start_col in admin_offsets:
            adm_pages: dict[str, dict[str, float]] = {}
            for p_idx, p_name in enumerate(ADMIN_PAGES):
                c0 = start_col + p_idx * 5
                p_amt = _parse_float(cell(r, c0))
                p_cust = _parse_float(cell(r, c0 + 1))
                d_amt = _parse_float(cell(r, c0 + 2))
                d_cust = _parse_float(cell(r, c0 + 3))
                adm_pages[p_name] = {
                    "print_amt": p_amt,
                    "print_cust": int(p_cust),
                    "dep_amt": d_amt,
                    "dep_cust": int(d_cust),
                }
            day_admins[adm] = adm_pages
        # มาย (col 124 if any)
        mai_val = _parse_float(cell(r, 124))
        day_admins["มาย"] = {
            p_name: {
                "print_amt": mai_val if p_idx == 0 else 0.0,
                "print_cust": 1 if (p_idx == 0 and mai_val > 0) else 0,
                "dep_amt": 0.0,
                "dep_cust": 0,
            }
            for p_idx, p_name in enumerate(ADMIN_PAGES)
        }
        days_raw[d_str] = {"date": d_str, "admins": day_admins}

    # Parse Ad Spend table (rows 74..82)
    ad_spend: list[dict[str, Any]] = []
    for r_idx in range(74, min(85, len(rows))):
        r = rows[r_idx]
        p_name = cell(r, 1)
        if not p_name or p_name == "เพจ":
            continue
        ad_spend.append({
            "page": p_name,
            "daily_budget": _parse_float(cell(r, 2)),
            "boost_pct": cell(r, 3),
            "adjusted_budget": _parse_float(cell(r, 4)) or _parse_float(cell(r, 2)),
            "boost_date": _norm_date(cell(r, 5), 10) if cell(r, 5) else "",
            "monthly_total": _parse_float(cell(r, 6)),
        })

    payload = {
        "synced_at": dt.datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        "days": days_raw,
        "ad_spend": ad_spend,
    }
    ADMIN_CACHE_PATH.write_text(json.dumps(payload, ensure_ascii=False), encoding="utf-8")
    return payload


def get_cached_admin_sheet() -> dict[str, Any]:
    if ADMIN_CACHE_PATH.exists():
        try:
            return json.loads(ADMIN_CACHE_PATH.read_text(encoding="utf-8"))
        except Exception:
            pass
    try:
        return sync_admin_page_sheet()
    except Exception:
        return {"synced_at": "", "days": {}, "ad_spend": []}


def _map_sales_to_admin(sales_str: str) -> str:
    s = str(sales_str or "").strip()
    if "จูน" in s:
        return "จูน"
    if "เอม" in s:
        return "เอม"
    if "ปอนด์" in s:
        return "ปอนด์"
    if any(k in s for k in ("มาย", "มิ้น", "ยีน", "ST")):
        return "มาย"
    return "จูน" if s else ""


def _map_channel_to_page(channel_str: str) -> str:
    ch = str(channel_str or "").strip().lower()
    if not ch:
        return ""
    if "st" in ch or "มาย" in ch:
        return "STprint"
    if "plp" in ch:
        return "PLP"
    if "โรงงาน" in ch:
        return "โรงงานสติ๊กเกอร์25"
    if "ถูกดีจริง" in ch:
        return "ถูกดีจริง"
    if "ถูกเว้ย" in ch:
        return "ถูกเว้ย25"
    if "ฉลาก25" in ch:
        return "ฉลาก25"
    if "พิมพ์สวย" in ch:
        return "พิมพ์สวย"
    if "ปริ้นฉลาก" in ch:
        return "ปริ้นฉลาก"
    return "STprint"


def build_admin_board_state(store: dict[str, Any], orders: list[dict[str, Any]] | None = None) -> dict[str, Any]:
    raw = get_cached_admin_sheet()
    days_map: dict[str, Any] = {}

    # 1. Calculate daily Admin & Page stats 100% directly from main queue sheet (ลงคิวจากแอดมิน💰2026)
    auto_days: dict[str, dict[str, dict[str, dict[str, float]]]] = {}

    def _ensure_auto_cell(d_s: str, adm_s: str, pg_s: str) -> dict[str, float]:
        d_rec = auto_days.setdefault(d_s, {})
        a_rec = d_rec.setdefault(adm_s, {
            p: {"print_amt": 0.0, "print_cust": 0, "dep_amt": 0.0, "dep_cust": 0}
            for p in ADMIN_PAGES
        })
        return a_rec.setdefault(pg_s, {"print_amt": 0.0, "print_cust": 0, "dep_amt": 0.0, "dep_cust": 0})

    for o in (orders or []):
        f_amt = float(o.get("full_amt") or 0.0)
        d_amt = float(o.get("dep_amt") or 0.0)
        if f_amt <= 0 and d_amt <= 0:
            continue
        adm = _map_sales_to_admin(str(o.get("sales") or "")) or "จูน"
        pg = _map_channel_to_page(str(o.get("channel") or "")) or "STprint"
        is_sub = bool(o.get("is_sub_row"))
        f_date = str(o.get("full_date") or o.get("date") or "")[:10]
        d_date = str(o.get("dep_date") or o.get("date") or "")[:10]

        if f_amt > 0 and f_date:
            c = _ensure_auto_cell(f_date, adm, pg)
            c["print_amt"] += f_amt
            if not is_sub:
                c["print_cust"] += 1

        if d_amt > 0 and d_date:
            c = _ensure_auto_cell(d_date, adm, pg)
            c["dep_amt"] += d_amt
            if not is_sub:
                c["dep_cust"] += 1

    for d_str, adm_dict in auto_days.items():
        day_rec = days_map.setdefault(d_str, {"date": d_str, "admins": {}})
        for adm, pg_dict in adm_dict.items():
            adm_rec = day_rec["admins"].setdefault(adm, {})
            for p_name, auto_v in pg_dict.items():
                adm_rec[p_name] = {
                    "print_amt": round(float(auto_v["print_amt"]), 2),
                    "print_cust": int(auto_v["print_cust"]),
                    "dep_amt": round(float(auto_v["dep_amt"]), 2),
                    "dep_cust": int(auto_v["dep_cust"]),
                }

    overrides: dict[str, Any] = store.get("admin_daily_overrides") or {}

    # Merge manual overrides
    for d_str, adm_map in overrides.items():
        day_rec = days_map.setdefault(d_str, {"date": d_str, "admins": {}})
        for adm, pages_override in (adm_map or {}).items():
            adm_rec = day_rec["admins"].setdefault(adm, {})
            for p_name, vals in (pages_override or {}).items():
                adm_rec[p_name] = {
                    "print_amt": _parse_float(vals.get("print_amt")),
                    "print_cust": int(_parse_float(vals.get("print_cust"))),
                    "dep_amt": _parse_float(vals.get("dep_amt")),
                    "dep_cust": int(_parse_float(vals.get("dep_cust"))),
                }

    # Build sorted daily list with candlestick & stacked bar metrics
    daily_list: list[dict[str, Any]] = []
    admin_totals: dict[str, dict[str, float]] = {
        a: {"print_amt": 0.0, "print_cust": 0, "dep_amt": 0.0, "dep_cust": 0, "total_amt": 0.0}
        for a in ADMIN_NAMES
    }
    page_totals: dict[str, dict[str, float]] = {
        p: {"print_amt": 0.0, "print_cust": 0, "dep_amt": 0.0, "dep_cust": 0, "total_amt": 0.0}
        for p in ADMIN_PAGES
    }

    prev_close = 0.0
    for d_str in sorted(days_map.keys()):
        d_obj = days_map[d_str]
        admins_data = d_obj.get("admins") or {}
        by_admin: dict[str, dict[str, float]] = {}
        by_page: dict[str, dict[str, float]] = {
            p: {"print_amt": 0.0, "print_cust": 0, "dep_amt": 0.0, "dep_cust": 0, "total_amt": 0.0}
            for p in ADMIN_PAGES
        }
        day_print = 0.0
        day_dep = 0.0
        day_print_cust = 0
        day_dep_cust = 0

        for adm in ADMIN_NAMES:
            p_map = admins_data.get(adm) or {}
            a_p_amt = 0.0
            a_p_cust = 0
            a_d_amt = 0.0
            a_d_cust = 0
            for p_name in ADMIN_PAGES:
                cell_v = p_map.get(p_name) or {}
                pa = float(cell_v.get("print_amt") or 0.0)
                pc = int(cell_v.get("print_cust") or 0)
                da = float(cell_v.get("dep_amt") or 0.0)
                dc = int(cell_v.get("dep_cust") or 0)
                a_p_amt += pa
                a_p_cust += pc
                a_d_amt += da
                a_d_cust += dc
                by_page[p_name]["print_amt"] += pa
                by_page[p_name]["print_cust"] += pc
                by_page[p_name]["dep_amt"] += da
                by_page[p_name]["dep_cust"] += dc
                by_page[p_name]["total_amt"] += pa + da
                page_totals[p_name]["print_amt"] += pa
                page_totals[p_name]["print_cust"] += pc
                page_totals[p_name]["dep_amt"] += da
                page_totals[p_name]["dep_cust"] += dc
                page_totals[p_name]["total_amt"] += pa + da

            a_tot = a_p_amt + a_d_amt
            by_admin[adm] = {
                "print_amt": a_p_amt,
                "print_cust": a_p_cust,
                "dep_amt": a_d_amt,
                "dep_cust": a_d_cust,
                "total_amt": a_tot,
                "pages": p_map,
            }
            admin_totals[adm]["print_amt"] += a_p_amt
            admin_totals[adm]["print_cust"] += a_p_cust
            admin_totals[adm]["dep_amt"] += a_d_amt
            admin_totals[adm]["dep_cust"] += a_d_cust
            admin_totals[adm]["total_amt"] += a_tot
            day_print += a_p_amt
            day_dep += a_d_amt
            day_print_cust += a_p_cust
            day_dep_cust += a_d_cust

        day_total = day_print + day_dep
        open_v = prev_close if prev_close > 0 else (day_print if day_total > 0 else 0.0)
        close_v = day_total
        high_v = max(open_v, close_v) + (day_dep if day_total > 0 else 0.0)
        low_v = min(open_v, close_v, day_print) if day_total > 0 else 0.0
        bullish = close_v >= open_v
        if day_total > 0:
            prev_close = close_v

        daily_list.append({
            "date": d_str,
            "print_amt": day_print,
            "print_cust": day_print_cust,
            "dep_amt": day_dep,
            "dep_cust": day_dep_cust,
            "total_amt": day_total,
            "open": open_v,
            "close": close_v,
            "high": high_v,
            "low": low_v,
            "bullish": bullish,
            "by_admin": by_admin,
            "by_page": by_page,
        })

    ad_spend_list = list(raw.get("ad_spend") or [])
    ad_overrides = store.get("ad_spend_overrides") or {}
    active_dates = [d["date"] for d in daily_list if d["total_amt"] > 0]
    if not active_dates:
        active_dates = [f"2026-10-{day:02d}" for day in range(1, dt.date.today().day + 1)]
    for item in ad_spend_list:
        ov = ad_overrides.get(item["page"])
        if ov:
            item.update(ov)
        if not item.get("monthly_total"):
            base_b = float(item.get("daily_budget") or 0.0)
            adj_b = float(item.get("adjusted_budget") or base_b)
            b_date = str(item.get("boost_date") or "").strip()
            tot_ad = 0.0
            for d_s in active_dates:
                if b_date and d_s >= b_date:
                    tot_ad += adj_b
                else:
                    tot_ad += base_b
            item["monthly_total"] = round(tot_ad, 2)

    return {
        "synced_at": raw.get("synced_at") or "",
        "sheet_id": GSHEET_ADMIN_ID,
        "pages": ADMIN_PAGES,
        "admins": ADMIN_NAMES,
        "daily": daily_list,
        "admin_totals": admin_totals,
        "page_totals": page_totals,
        "ad_spend": ad_spend_list,
    }


def _ws_eval_js_cdp(ws_url: str, expression: str, timeout: float = 4.0) -> Any:
    parsed = urllib.parse.urlparse(ws_url)
    host = parsed.hostname or "127.0.0.1"
    port = parsed.port or MYORDER_CDP_PORT
    path = parsed.path or "/"
    sock = socket.create_connection((host, port), timeout=timeout)
    try:
        key = base64.b64encode(os.urandom(16)).decode("ascii")
        handshake = (
            f"GET {path} HTTP/1.1\r\n"
            f"Host: {host}:{port}\r\n"
            "Upgrade: websocket\r\n"
            "Connection: Upgrade\r\n"
            f"Sec-WebSocket-Key: {key}\r\n"
            "Sec-WebSocket-Version: 13\r\n\r\n"
        )
        sock.sendall(handshake.encode("ascii"))
        buf = b""
        while b"\r\n\r\n" not in buf:
            chunk = sock.recv(4096)
            if not chunk:
                return None
            buf += chunk
        leftover = buf[buf.index(b"\r\n\r\n") + 4 :]
        payload_bytes = json.dumps(
            {
                "id": 1,
                "method": "Runtime.evaluate",
                "params": {"expression": expression, "returnByValue": True, "awaitPromise": True},
            }
        ).encode("utf-8")
        mask_key = os.urandom(4)
        masked = bytes(b ^ mask_key[i % 4] for i, b in enumerate(payload_bytes))
        length = len(payload_bytes)
        if length < 126:
            header = struct.pack("!BB", 0x81, 0x80 | length)
        elif length < 65536:
            header = struct.pack("!BBH", 0x81, 0x80 | 126, length)
        else:
            header = struct.pack("!BBQ", 0x81, 0x80 | 127, length)
        sock.sendall(header + mask_key + masked)

        def recv_exact(n: int) -> bytes:
            nonlocal leftover
            while len(leftover) < n:
                c = sock.recv(65536)
                if not c:
                    raise RuntimeError("EOF")
                leftover += c
            out = leftover[:n]
            leftover = leftover[n:]
            return out

        while True:
            b1, b2 = recv_exact(2)
            opcode = b1 & 0x0F
            is_masked = bool(b2 & 0x80)
            flen = b2 & 0x7F
            if flen == 126:
                (flen,) = struct.unpack("!H", recv_exact(2))
            elif flen == 127:
                (flen,) = struct.unpack("!Q", recv_exact(8))
            mkey = recv_exact(4) if is_masked else b""
            fdata = recv_exact(flen)
            if is_masked:
                fdata = bytes(b ^ mkey[i % 4] for i, b in enumerate(fdata))
            if opcode == 0x8:
                return None
            if opcode == 0x9:
                sock.sendall(struct.pack("!BB", 0x8A, 0x80) + os.urandom(4))
                continue
            if opcode == 0x1:
                msg = json.loads(fdata.decode("utf-8", errors="replace"))
                if msg.get("id") == 1:
                    return ((msg.get("result") or {}).get("result") or {}).get("value")
    finally:
        try:
            sock.close()
        except Exception:
            pass


def _get_myorder_access_token() -> str:
    # 1. Try reading live accessToken from Chrome CDP (port 9225)
    try:
        req = urllib.request.Request(f"http://127.0.0.1:{MYORDER_CDP_PORT}/json")
        with urllib.request.urlopen(req, timeout=2.0) as resp:
            tabs = json.loads(resp.read().decode("utf-8"))
        ws_url = next(
            (
                t.get("webSocketDebuggerUrl")
                for t in tabs
                if t.get("type") == "page" and "myorder.ai" in str(t.get("url") or "") and t.get("webSocketDebuggerUrl")
            ),
            None,
        )
        if ws_url:
            tok = str(_ws_eval_js_cdp(ws_url, "localStorage.getItem('accessToken') || ''") or "").strip()
            if tok:
                try:
                    MYORDER_TOKEN_PATH.write_text(tok, encoding="utf-8")
                except Exception:
                    pass
                return tok
    except Exception:
        pass

    # 2. Fallback to cached token file
    if MYORDER_TOKEN_PATH.exists():
        try:
            tok = MYORDER_TOKEN_PATH.read_text(encoding="utf-8").strip()
            if tok:
                return tok
        except Exception:
            pass

    # 3. Fallback to scanning myorder-bot Chrome Local Storage LevelDB
    try:
        ldb_dir = BASE_DIR.parent / "myorder-bot" / "data" / "chrome-profile" / "Default" / "Local Storage" / "leveldb"
        if ldb_dir.exists():
            files = sorted(ldb_dir.glob("*.*"), key=lambda f: f.stat().st_mtime, reverse=True)
            for f in files:
                if f.suffix.lower() not in (".log", ".ldb"):
                    continue
                raw = f.read_bytes()
                hits = re.findall(rb"eyJ[A-Za-z0-9_-]{20,}\.eyJ[A-Za-z0-9_-]{20,}\.[A-Za-z0-9_-]{20,}", raw)
                if hits:
                    tok = hits[-1].decode("ascii", errors="ignore")
                    MYORDER_TOKEN_PATH.write_text(tok, encoding="utf-8")
                    return tok
    except Exception:
        pass
    return ""


def _map_myorder_carrier(raw_express: str) -> str:
    s = str(raw_express or "").strip().upper()
    if "THAI_POST" in s or "THAIPOST" in s or "EMS" in s:
        return "ไปรษณีย์ไทย"
    if "FLASH" in s:
        return "Flash Express"
    if "J&T" in s or "JANDT" in s or "JT" in s:
        return "J&T Express"
    if "KERRY" in s or "KEX" in s:
        return "Kerry Express"
    if "SHOPEE" in s or "SPX" in s:
        return "SPX Express"
    if "DHL" in s:
        return "DHL"
    if "NINJA" in s:
        return "Ninja Van"
    return raw_express or "ไปรษณีย์ไทย"


def _normalize_myorder_record(x: dict[str, Any], is_pending: bool = False) -> dict[str, Any]:
    cd = x.get("customerData") or {}
    ex = x.get("express") or {}
    raw_date = str(x.get("date") or x.get("createdAt") or "").strip()
    bkk_date = ""
    if raw_date:
        try:
            dt_utc = dt.datetime.fromisoformat(raw_date.replace("Z", "+00:00"))
            bkk_date = (dt_utc + dt.timedelta(hours=7)).strftime("%Y-%m-%d")
        except Exception:
            bkk_date = raw_date[:10]

    tel_digits = re.sub(r"\D+", "", str(cd.get("tel") or ""))
    if tel_digits.startswith("66") and len(tel_digits) >= 11:
        tel_digits = "0" + tel_digits[2:]
    tel = tel_digits[:10] if len(tel_digits) >= 9 else ""

    cname = str(cd.get("customerName") or "").strip()
    comment = str(cd.get("comment") or "").strip()
    tracking_no = str(ex.get("trackingNumber") or "").strip()
    carrier = _map_myorder_carrier(str(ex.get("expressName") or "")) if (ex.get("expressName") or tracking_no) else ""
    is_printed = bool(ex.get("isPrinted"))
    picked_up = bool(ex.get("pickedUpTimestamp"))

    st_obj = ex.get("shippingState") or {}
    state_code = str(st_obj.get("state") or "").strip()
    state_text = str(st_obj.get("stateText") or "").strip()

    loc_list = ex.get("shippingLocalStatus") or []
    last_step = ""
    last_step_time = ""
    if isinstance(loc_list, list) and loc_list:
        last_item = loc_list[-1] if isinstance(loc_list[-1], dict) else {}
        last_step = str(last_item.get("stateText") or "").strip()
        raw_ep = last_item.get("stateDate")
        if isinstance(raw_ep, (int, float)) and raw_ep > 0:
            try:
                ep_sec = float(raw_ep) / 1000.0 if raw_ep > 1e11 else float(raw_ep)
                dt_step = dt.datetime.fromtimestamp(ep_sec, dt.timezone(dt.timedelta(hours=7)))
                last_step_time = dt_step.strftime("%d/%m %H:%M")
            except Exception:
                pass

    if is_pending and not tracking_no:
        shipping_status = "ready_to_pack"
        myorder_badge = "📦 เปิดบิล MyOrder (รอออกเลข)"
    elif state_code in ("4", "6", "7") or "รับพัสดุ" in state_text or "นำจ่ายถึงผู้รับแล้ว" in last_step:
        shipping_status = "delivered"
        myorder_badge = "✅ รับพัสดุแล้ว (MyOrder)"
    elif state_code in ("5", "8", "9", "10", "11", "12") or any(
        w in state_text for w in ("ปัญหา", "ตีกลับ", "ไม่สำเร็จ", "ยกเลิก", "สูญหาย", "เสียหาย")
    ):
        shipping_status = "issue"
        myorder_badge = f"⚠️ {state_text or 'พัสดุมีปัญหา'} (MyOrder)"
    elif state_code == "3" or "ปลายทาง" in state_text or "ปลายทาง" in last_step:
        shipping_status = "shipped"
        myorder_badge = "🚚 ถึงสาขาปลายทาง (MyOrder)"
    elif state_code in ("1", "2") or last_step or picked_up:
        shipping_status = "shipped"
        myorder_badge = f"🚚 {state_text or 'กำลังขนส่ง'} (MyOrder)"
    elif tracking_no:
        shipping_status = "shipped" if is_printed else "packed"
        myorder_badge = "🖨️ พิมพ์ใบส่งของแล้ว (MyOrder)" if is_printed else "🏷️ ออกเลขพัสดุแล้ว (MyOrder)"
    else:
        shipping_status = "ready_to_pack"
        myorder_badge = "📦 รอออกเลขพัสดุ (MyOrder)"

    return {
        "myorder_id": str(x.get("_id") or ""),
        "date": bkk_date,
        "customer_name": cname,
        "tel": tel,
        "comment": comment,
        "carrier": carrier,
        "tracking_no": tracking_no,
        "is_printed": is_printed,
        "state_code": state_code,
        "state_text": state_text,
        "last_step": last_step,
        "last_step_time": last_step_time,
        "shipping_status": shipping_status,
        "myorder_badge": myorder_badge,
    }


def sync_myorder_shipping() -> dict[str, Any]:
    """Fetch complete & pending orders with tracking numbers and live courier states from MyOrder API."""
    tok = _get_myorder_access_token()
    if not tok:
        raise RuntimeError("ไม่พบโทเคนล็อกอิน MyOrder (กรุณาเปิดหน้า MyOrder ใน Chrome พอร์ต 9225)")

    headers = {
        "Authorization": f"Bearer {tok}",
        "Content-Type": "application/json",
        "Accept": "application/json, text/plain, */*",
        "Origin": "https://www.myorder.ai",
        "Referer": "https://www.myorder.ai/",
        "User-Agent": MYORDER_USER_AGENT,
    }

    normalized_orders: list[dict[str, Any]] = []
    seen_ids: set[str] = set()

    # 1. Fetch pending orders (คำสั่งซื้อรอดำเนินการ)
    try:
        req_p = urllib.request.Request(
            f"{MYORDER_OM_BASE}/order-management/pending/filter",
            data=json.dumps({"teamId": MYORDER_TEAM_ID, "type": "ALL", "page": 1, "limit": 200}).encode("utf-8"),
            method="POST",
            headers=headers,
        )
        with urllib.request.urlopen(req_p, timeout=15) as r:
            dp = json.loads(r.read().decode("utf-8"))
            rows_p = (dp.get("data") or {}).get("datas") or dp.get("datas") or []
            for x in rows_p:
                rec = _normalize_myorder_record(x, is_pending=True)
                mid = rec["myorder_id"]
                if mid and mid not in seen_ids:
                    seen_ids.add(mid)
                    normalized_orders.append(rec)
    except Exception:
        pass

    # 2. Fetch completed/shipped orders (คำสั่งซื้อรอจัดส่ง / ออกเลขพัสดุแล้ว)
    start_dt = (dt.date.today() - dt.timedelta(days=45)).isoformat() + "T00:00:00.000Z"
    end_dt = (dt.date.today() + dt.timedelta(days=2)).isoformat() + "T23:59:59.999Z"
    for pg in (1, 2, 3):
        payload = {
            "teamId": MYORDER_TEAM_ID,
            "type": "ALL",
            "page": pg,
            "limit": 800,
            "startDate": start_dt,
            "endDate": end_dt,
        }
        req_c = urllib.request.Request(
            f"{MYORDER_OM_BASE}/order-management/complete/filter",
            data=json.dumps(payload).encode("utf-8"),
            method="POST",
            headers=headers,
        )
        with urllib.request.urlopen(req_c, timeout=20) as r:
            dc = json.loads(r.read().decode("utf-8"))
            rows_c = (dc.get("data") or {}).get("datas") or dc.get("datas") or []
            if not rows_c:
                break
            for x in rows_c:
                rec = _normalize_myorder_record(x, is_pending=False)
                mid = rec["myorder_id"]
                if mid and mid not in seen_ids:
                    seen_ids.add(mid)
                    normalized_orders.append(rec)
            if len(rows_c) < 800:
                break

    cache_payload = {
        "synced_at": dt.datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        "orders": normalized_orders,
    }
    MYORDER_CACHE_PATH.write_text(json.dumps(cache_payload, ensure_ascii=False), encoding="utf-8")
    return cache_payload


def get_cached_myorder_shipping() -> dict[str, Any]:
    if MYORDER_CACHE_PATH.exists():
        try:
            return json.loads(MYORDER_CACHE_PATH.read_text(encoding="utf-8"))
        except Exception:
            pass
    try:
        return sync_myorder_shipping()
    except Exception as e:
        return {"synced_at": f"ยังไม่เชื่อมต่อ ({e})", "orders": []}


def _clean_match_str(s: str) -> str:
    return re.sub(r"[\s\-_•🔹()+]+", "", str(s or "")).lower()


def get_cached_sheet_data() -> dict[str, Any]:
    if SHEET_CACHE_PATH.exists():
        try:
            return json.loads(SHEET_CACHE_PATH.read_text(encoding="utf-8"))
        except Exception:
            pass
    try:
        return sync_from_google_sheet(all_months=False)
    except Exception as e:
        return {"synced_at": f"ดึงข้อมูลไม่สำเร็จ ({e})", "orders": [], "incomes": []}


def build_full_state() -> dict[str, Any]:
    with _LOCK:
        store = _load_store()
    sheet_data = get_cached_sheet_data()
    myorder_data = get_cached_myorder_shipping()
    overrides = store.get("order_overrides") or {}

    # Build fast lookup indices for MyOrder orders
    mo_list = myorder_data.get("orders") or []
    mo_by_tel: dict[str, list[dict[str, Any]]] = {}
    mo_by_name: dict[str, list[dict[str, Any]]] = {}
    mo_by_job: dict[str, list[dict[str, Any]]] = {}
    for m in mo_list:
        t = str(m.get("tel") or "")
        if len(t) >= 9:
            mo_by_tel.setdefault(t[:10], []).append(m)
        cn = _clean_match_str(m.get("customer_name") or "")
        if len(cn) >= 2:
            mo_by_name.setdefault(cn, []).append(m)
        cm = _clean_match_str(m.get("comment") or "")
        if len(cm) >= 3:
            mo_by_job.setdefault(cm, []).append(m)

    def find_best_myorder(o_rec: dict[str, Any]) -> dict[str, Any] | None:
        # Deposit-only row without full payment, phone, or printed status should not prematurely claim a shipment
        is_deposit_only = (
            float(o_rec.get("dep_amt") or 0) > 0
            and float(o_rec.get("full_amt") or 0) <= 0
            and not o_rec.get("is_sub_row")
            and not o_rec.get("phone")
            and o_rec.get("prep_status_code") != "printed"
        )
        if is_deposit_only:
            return None

        tel_d = re.sub(r"\D+", "", str(o_rec.get("phone") or ""))
        if tel_d.startswith("66") and len(tel_d) >= 11:
            tel_d = "0" + tel_d[2:]
        tel_key = tel_d[:10] if len(tel_d) >= 9 else ""
        rec_key = _clean_match_str(o_rec.get("recipient") or "")
        fb_key = _clean_match_str(o_rec.get("fb") or "")
        job_key = _clean_match_str(o_rec.get("job") or "")

        candidates: list[dict[str, Any]] = []
        seen_c: set[str] = set()
        for src in (
            mo_by_tel.get(tel_key, []) if tel_key else [],
            mo_by_name.get(rec_key, []) if len(rec_key) >= 2 else [],
            mo_by_name.get(fb_key, []) if len(fb_key) >= 2 else [],
            mo_by_job.get(job_key, []) if len(job_key) >= 3 else [],
        ):
            for c in src:
                cid = c.get("myorder_id") or id(c)
                if cid not in seen_c:
                    seen_c.add(cid)
                    candidates.append(c)
        if not candidates:
            return None

        o_dt: dt.date | None = None
        try:
            o_dt = dt.date.fromisoformat(str(o_rec.get("date") or "")[:10])
        except Exception:
            pass

        best_c: dict[str, Any] | None = None
        best_score = -1e9
        for c in candidates:
            score = 0.0
            if tel_key and c.get("tel") == tel_key:
                score += 100.0
            c_name = _clean_match_str(c.get("customer_name") or "")
            if rec_key and c_name and (rec_key == c_name or rec_key in c_name or c_name in rec_key):
                score += 60.0
            elif fb_key and c_name and (fb_key == c_name or fb_key in c_name or c_name in fb_key):
                score += 50.0
            c_job = _clean_match_str(c.get("comment") or "")
            if job_key and c_job and (job_key == c_job or job_key in c_job or c_job in job_key):
                score += 80.0
            if o_dt and c.get("date"):
                try:
                    c_dt = dt.date.fromisoformat(str(c["date"])[:10])
                    diff_days = (c_dt - o_dt).days
                    # Usually shipped 0 to 5 days after order date
                    if -2 <= diff_days <= 7:
                        score += 40.0 - abs(diff_days) * 3.0
                    else:
                        score -= abs(diff_days) * 4.0
                except Exception:
                    pass
            if c.get("tracking_no"):
                score += 10.0
            if score > best_score:
                best_score = score
                best_c = c
        return best_c if best_score >= 40.0 else None

    orders: list[dict[str, Any]] = []
    myorder_matched_count = 0
    for o in sheet_data.get("orders", []):
        oc = dict(o)
        ov = overrides.get(oc["id"]) or {}
        mo = find_best_myorder(oc)
        if mo:
            myorder_matched_count += 1
            oc["myorder_matched"] = True
            oc["myorder_id"] = mo.get("myorder_id") or ""
            oc["myorder_badge"] = mo.get("myorder_badge") or ""
            oc["myorder_state_text"] = mo.get("state_text") or ""
            oc["myorder_step"] = mo.get("last_step") or ""
            oc["myorder_step_time"] = mo.get("last_step_time") or ""
            default_ship_status = mo.get("shipping_status") or "shipped"
            default_tracking = mo.get("tracking_no") or ""
            default_carrier = mo.get("carrier") or ""
            if not oc.get("phone") and mo.get("tel"):
                oc["phone"] = mo["tel"]
            if not oc.get("recipient") and mo.get("customer_name"):
                oc["recipient"] = mo["customer_name"]
        else:
            oc["myorder_matched"] = False
            oc["myorder_id"] = ""
            oc["myorder_badge"] = ""
            oc["myorder_state_text"] = ""
            oc["myorder_step"] = ""
            oc["myorder_step_time"] = ""
            default_ship_status = "ready_to_pack" if oc["prep_status_code"] == "printed" else "waiting_prod"
            default_tracking = ""
            default_carrier = ""

        oc["shipping_status"] = ov.get("shipping_status") or default_ship_status
        oc["tracking_no"] = ov.get("tracking_no") or default_tracking
        oc["carrier"] = ov.get("carrier") or default_carrier
        if ov.get("prep_status_override"):
            oc["prep_status_code"] = ov["prep_status_override"]
            oc["prep_status_label"] = ov.get("prep_status_label") or oc["prep_status_label"]
        if ov.get("admin_note"):
            oc["admin_note"] = ov["admin_note"]
        else:
            oc["admin_note"] = ""
        orders.append(oc)

    # Combine sheet transfer incomes + manual incomes
    manual_incomes = store.get("manual_incomes") or []
    all_incomes = list(sheet_data.get("incomes", [])) + list(manual_incomes)
    all_incomes.sort(key=lambda x: (str(x.get("date") or ""), str(x.get("time") or ""), int(x.get("row") or 0)), reverse=True)

    # Calculate stock usage from orders
    mat_sheets_used: dict[str, float] = {}
    primary_orders_count = 0
    for o in orders:
        if not o.get("is_sub_row") and (o.get("total_amt", 0) > 0 or o.get("total_sheets", 0) > 0):
            primary_orders_count += 1
        mat = str(o.get("material") or "").strip()
        sh = float(o.get("total_sheets") or 0.0)
        if mat and sh > 0:
            mat_sheets_used[mat] = mat_sheets_used.get(mat, 0.0) + sh

    movements = store.get("stock_movements") or []
    mov_by_item: dict[str, dict[str, float]] = {}
    for m in movements:
        iid = str(m.get("item_id") or "")
        mtype = str(m.get("type") or "in")
        qty = float(m.get("qty") or 0.0)
        rec = mov_by_item.setdefault(iid, {"in": 0.0, "out": 0.0})
        if mtype == "in":
            rec["in"] += qty
        else:
            rec["out"] += qty

    stock_items_computed: list[dict[str, Any]] = []
    for item in store.get("stock_items") or []:
        ic = dict(item)
        iid = ic["id"]
        init_q = float(ic.get("initial_qty") or 0.0)
        mv = mov_by_item.get(iid, {"in": 0.0, "out": 0.0})
        auto_used = 0.0
        for m_match in ic.get("material_match") or []:
            if m_match == "__ORDER_COUNT__":
                auto_used += float(primary_orders_count)
            else:
                auto_used += float(mat_sheets_used.get(m_match, 0.0))
        total_in = init_q + mv["in"]
        total_out = auto_used + mv["out"]
        remain = total_in - total_out
        ic["stock_in_added"] = mv["in"]
        ic["manual_out"] = mv["out"]
        ic["auto_used_from_orders"] = auto_used
        ic["total_used"] = total_out
        ic["remaining_qty"] = remain
        ic["low_stock"] = remain <= float(ic.get("min_qty") or 0.0)
        stock_items_computed.append(ic)

    # Build Customer CRM aggregation
    cust_map: dict[str, dict[str, Any]] = {}
    for o in orders:
        ckey = o.get("customer_key") or "unknown"
        if ckey == "unknown" and not o.get("fb"):
            continue
        c = cust_map.get(ckey)
        if c is None:
            c = {
                "customer_key": ckey,
                "fb": o.get("fb") or o.get("recipient") or "ไม่ระบุชื่อ",
                "recipient": o.get("recipient") or "",
                "phone": o.get("phone") or "",
                "address": o.get("address") or "",
                "channel": o.get("channel") or "",
                "sales": o.get("sales") or "",
                "order_count": 0,
                "design_count": 0,
                "total_deposit": 0.0,
                "total_full": 0.0,
                "total_spent": 0.0,
                "total_sheets": 0.0,
                "last_order_date": "",
                "orders": [],
            }
            cust_map[ckey] = c
        if o.get("fb") and (not c["fb"] or c["fb"] == "ไม่ระบุชื่อ"):
            c["fb"] = o["fb"]
        if o.get("recipient") and not c["recipient"]:
            c["recipient"] = o["recipient"]
        if o.get("phone") and not c["phone"]:
            c["phone"] = o["phone"]
        if o.get("address") and not c["address"]:
            c["address"] = o["address"]
        if o.get("channel") and not c["channel"]:
            c["channel"] = o["channel"]
        if not o.get("is_sub_row"):
            c["order_count"] += 1
        c["design_count"] += 1
        c["total_deposit"] += float(o.get("dep_amt") or 0.0)
        c["total_full"] += float(o.get("full_amt") or 0.0)
        c["total_spent"] += float(o.get("total_amt") or 0.0)
        c["total_sheets"] += float(o.get("total_sheets") or 0.0)
        if str(o.get("date") or "") > c["last_order_date"]:
            c["last_order_date"] = str(o.get("date") or "")
        c["orders"].append(o)

    customers = sorted(cust_map.values(), key=lambda x: (x["last_order_date"], x["total_spent"]), reverse=True)
    admin_board = build_admin_board_state(store, orders=orders)

    return {
        "synced_at": sheet_data.get("synced_at") or "",
        "myorder_info": {
            "synced_at": myorder_data.get("synced_at") or "",
            "total_orders": len(mo_list),
            "matched_count": myorder_matched_count,
        },
        "mobile_info": {
            "lan_url": f"http://{_get_lan_ip()}:{PORT}",
            "public_url": _PUBLIC_TUNNEL_URL,
            "permanent_url": PERMANENT_GATEWAY_URL,
        },
        "orders": orders,
        "incomes": all_incomes,
        "expenses": sorted(store.get("expenses") or [], key=lambda x: str(x.get("date") or ""), reverse=True),
        "stock_items": stock_items_computed,
        "stock_movements": sorted(movements, key=lambda x: str(x.get("ts") or ""), reverse=True),
        "customers": customers,
        "material_breakdown": mat_sheets_used,
        "admin_board": admin_board,
        "users": store.get("users") or DEFAULT_USERS,
    }


class DashboardHandler(BaseHTTPRequestHandler):
    def log_message(self, fmt: str, *args: Any) -> None:
        return

    def _send_cors_headers(self) -> None:
        self.send_header("Access-Control-Allow-Origin", "*")
        self.send_header("Access-Control-Allow-Methods", "GET, POST, OPTIONS")
        self.send_header("Access-Control-Allow-Headers", "Content-Type, Authorization")

    def do_OPTIONS(self) -> None:
        self.send_response(204)
        self._send_cors_headers()
        self.end_headers()

    def _send_json(self, data: Any, status: int = 200) -> None:
        raw = json.dumps(data, ensure_ascii=False).encode("utf-8")
        self.send_response(status)
        self.send_header("Content-Type", "application/json; charset=utf-8")
        self.send_header("Content-Length", str(len(raw)))
        self.send_header("Cache-Control", "no-store")
        self._send_cors_headers()
        self.end_headers()
        self.wfile.write(raw)

    def _read_body(self) -> dict[str, Any]:
        length = int(self.headers.get("Content-Length") or 0)
        if length <= 0:
            return {}
        raw = self.rfile.read(length)
        return json.loads(raw.decode("utf-8"))

    def do_GET(self) -> None:
        parsed = urllib.parse.urlparse(self.path)
        if parsed.path in ("/", "/index.html"):
            html_path = BASE_DIR / "index.html"
            raw = html_path.read_bytes()
            self.send_response(200)
            self.send_header("Content-Type", "text/html; charset=utf-8")
            self.send_header("Content-Length", str(len(raw)))
            self.send_header("Cache-Control", "no-store")
            self._send_cors_headers()
            self.end_headers()
            self.wfile.write(raw)
            return

        if parsed.path == "/api/ping":
            self._send_json({"ok": True, "ts": time.time()})
            return

        if parsed.path == "/api/state":
            self._send_json(build_full_state())
            return

        self._send_json({"error": "Not found"}, status=404)

    def do_POST(self) -> None:
        parsed = urllib.parse.urlparse(self.path)
        body = self._read_body()

        if parsed.path == "/api/sync":
            all_months = bool(body.get("all_months"))
            try:
                sync_from_google_sheet(all_months=all_months)
                try:
                    sync_admin_page_sheet()
                except Exception:
                    pass
                try:
                    sync_myorder_shipping()
                except Exception:
                    pass
                self._send_json({"ok": True, "state": build_full_state()})
            except Exception as e:
                self._send_json({"ok": False, "error": str(e)}, status=500)
            return

        if parsed.path == "/api/myorder/sync":
            try:
                sync_myorder_shipping()
                self._send_json({"ok": True, "state": build_full_state()})
            except Exception as e:
                self._send_json({"ok": False, "error": str(e)}, status=500)
            return

        if parsed.path == "/api/income/add":
            with _LOCK:
                store = _load_store()
                rec = {
                    "id": "inc-man-" + uuid.uuid4().hex[:8],
                    "source": "manual",
                    "transfer_type": str(body.get("transfer_type") or "other"),
                    "transfer_type_label": str(body.get("transfer_type_label") or "รายรับโอนอื่นๆ"),
                    "date": str(body.get("date") or dt.date.today().isoformat()),
                    "time": str(body.get("time") or dt.datetime.now().strftime("%H:%M")),
                    "amount": _parse_float(body.get("amount")),
                    "method": str(body.get("method") or "โอนเงิน"),
                    "verified": True,
                    "fb": str(body.get("fb") or "").strip(),
                    "recipient": str(body.get("recipient") or "").strip(),
                    "job": str(body.get("job") or "").strip(),
                    "channel": str(body.get("channel") or "").strip(),
                    "sales": str(body.get("sales") or "").strip(),
                    "note": str(body.get("note") or "").strip(),
                }
                store.setdefault("manual_incomes", []).append(rec)
                _save_store(store)
            self._send_json({"ok": True, "state": build_full_state()})
            return

        if parsed.path == "/api/income/delete":
            inc_id = str(body.get("id") or "")
            with _LOCK:
                store = _load_store()
                store["manual_incomes"] = [x for x in store.get("manual_incomes", []) if x.get("id") != inc_id]
                _save_store(store)
            self._send_json({"ok": True, "state": build_full_state()})
            return

        if parsed.path == "/api/expense/save":
            with _LOCK:
                store = _load_store()
                exps = store.setdefault("expenses", [])
                exp_id = str(body.get("id") or "").strip()
                rec = {
                    "id": exp_id or ("exp-" + uuid.uuid4().hex[:8]),
                    "date": str(body.get("date") or dt.date.today().isoformat()),
                    "due_date": str(body.get("due_date") or "").strip(),
                    "category": str(body.get("category") or "ค่าสติ๊กเกอร์/วัตถุดิบ").strip(),
                    "title": str(body.get("title") or "").strip(),
                    "vendor": str(body.get("vendor") or "").strip(),
                    "amount": _parse_float(body.get("amount")),
                    "paid_amount": _parse_float(body.get("paid_amount")),
                    "status": str(body.get("status") or "paid").strip(),  # paid | creditor | partial
                    "payment_method": str(body.get("payment_method") or "โอนเงิน").strip(),
                    "note": str(body.get("note") or "").strip(),
                }
                if rec["status"] == "paid":
                    rec["paid_amount"] = rec["amount"]
                elif rec["status"] == "creditor":
                    rec["paid_amount"] = 0.0
                idx = next((i for i, x in enumerate(exps) if x.get("id") == rec["id"]), None)
                if idx is not None:
                    exps[idx] = rec
                else:
                    exps.append(rec)
                _save_store(store)
            self._send_json({"ok": True, "state": build_full_state()})
            return

        if parsed.path == "/api/expense/mark_paid":
            exp_id = str(body.get("id") or "")
            with _LOCK:
                store = _load_store()
                for x in store.get("expenses", []):
                    if x.get("id") == exp_id:
                        x["status"] = "paid"
                        x["paid_amount"] = float(x.get("amount") or 0.0)
                _save_store(store)
            self._send_json({"ok": True, "state": build_full_state()})
            return

        if parsed.path == "/api/expense/delete":
            exp_id = str(body.get("id") or "")
            with _LOCK:
                store = _load_store()
                store["expenses"] = [x for x in store.get("expenses", []) if x.get("id") != exp_id]
                _save_store(store)
            self._send_json({"ok": True, "state": build_full_state()})
            return

        if parsed.path == "/api/stock/item/save":
            with _LOCK:
                store = _load_store()
                items = store.setdefault("stock_items", [])
                iid = str(body.get("id") or "").strip()
                existing = next((x for x in items if x.get("id") == iid), None)
                rec = {
                    "id": iid or ("stk-" + uuid.uuid4().hex[:8]),
                    "code": str(body.get("code") or "STK-NEW").strip(),
                    "name": str(body.get("name") or "").strip(),
                    "category": str(body.get("category") or "สติ๊กเกอร์").strip(),
                    "material_match": existing.get("material_match", []) if existing else (
                        [str(body.get("material_key")).strip()] if body.get("material_key") else []
                    ),
                    "unit": str(body.get("unit") or "แผ่น").strip(),
                    "initial_qty": _parse_float(body.get("initial_qty")),
                    "min_qty": _parse_float(body.get("min_qty")),
                    "cost_per_unit": _parse_float(body.get("cost_per_unit")),
                    "note": str(body.get("note") or "").strip(),
                }
                if existing:
                    idx = items.index(existing)
                    items[idx] = rec
                else:
                    items.append(rec)
                _save_store(store)
            self._send_json({"ok": True, "state": build_full_state()})
            return

        if parsed.path == "/api/stock/movement":
            mtype = str(body.get("type") or "in")
            iid = str(body.get("item_id") or "")
            qty = _parse_float(body.get("qty"))
            if mtype == "set_balance":
                curr_state = build_full_state()
                curr_item = next((x for x in curr_state.get("stock_items", []) if x.get("id") == iid), None)
                with _LOCK:
                    store = _load_store()
                    item = next((x for x in store.get("stock_items", []) if x.get("id") == iid), None)
                    if item and curr_item:
                        auto_u = float(curr_item.get("auto_used_from_orders") or 0.0)
                        mv_in = float(curr_item.get("stock_in_added") or 0.0)
                        mv_out = float(curr_item.get("manual_out") or 0.0)
                        item["initial_qty"] = qty + auto_u + mv_out - mv_in
                        mov = {
                            "id": "mov-" + uuid.uuid4().hex[:8],
                            "ts": dt.datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                            "date": str(body.get("date") or dt.date.today().isoformat()),
                            "item_id": iid,
                            "item_name": item.get("name") or iid,
                            "type": "set_balance",
                            "qty": 0.0,
                            "note": f"ตั้งยอดคงเหลือจริงตามนับสต๊อก = {int(qty):,} {item.get('unit','')} ({str(body.get('note') or '').strip()})".strip(" ()"),
                        }
                        store.setdefault("stock_movements", []).append(mov)
                        _save_store(store)
            else:
                with _LOCK:
                    store = _load_store()
                    item = next((x for x in store.get("stock_items", []) if x.get("id") == iid), None)
                    mov = {
                        "id": "mov-" + uuid.uuid4().hex[:8],
                        "ts": dt.datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                        "date": str(body.get("date") or dt.date.today().isoformat()),
                        "item_id": iid,
                        "item_name": item.get("name") if item else iid,
                        "type": mtype,  # in | out
                        "qty": qty,
                        "note": str(body.get("note") or "").strip(),
                    }
                    store.setdefault("stock_movements", []).append(mov)
                    _save_store(store)
            self._send_json({"ok": True, "state": build_full_state()})
            return

        if parsed.path == "/api/order/update":
            order_id = str(body.get("id") or "").strip()
            if order_id:
                with _LOCK:
                    store = _load_store()
                    ov = store.setdefault("order_overrides", {}).setdefault(order_id, {})
                    if "shipping_status" in body:
                        ov["shipping_status"] = str(body["shipping_status"])
                    if "tracking_no" in body:
                        ov["tracking_no"] = str(body["tracking_no"]).strip()
                    if "carrier" in body:
                        ov["carrier"] = str(body["carrier"]).strip()
                    if "prep_status_override" in body:
                        ov["prep_status_override"] = str(body["prep_status_override"]).strip()
                    if "prep_status_label" in body:
                        ov["prep_status_label"] = str(body["prep_status_label"]).strip()
                    if "admin_note" in body:
                        ov["admin_note"] = str(body["admin_note"]).strip()
                    _save_store(store)
            self._send_json({"ok": True, "state": build_full_state()})
            return

        if parsed.path == "/api/admin_board/save_day":
            d_str = str(body.get("date") or dt.date.today().isoformat()).strip()
            adm = str(body.get("admin") or "จูน").strip()
            pages_payload = body.get("pages") or {}
            with _LOCK:
                store = _load_store()
                day_ov = store.setdefault("admin_daily_overrides", {}).setdefault(d_str, {}).setdefault(adm, {})
                for p_name, vals in pages_payload.items():
                    day_ov[p_name] = {
                        "print_amt": _parse_float(vals.get("print_amt")),
                        "print_cust": int(_parse_float(vals.get("print_cust"))),
                        "dep_amt": _parse_float(vals.get("dep_amt")),
                        "dep_cust": int(_parse_float(vals.get("dep_cust"))),
                    }
                _save_store(store)
            self._send_json({"ok": True, "state": build_full_state()})
            return

        if parsed.path == "/api/admin_board/save_ads":
            ads_payload = body.get("ads") or []
            with _LOCK:
                store = _load_store()
                ad_ov = store.setdefault("ad_spend_overrides", {})
                for row in ads_payload:
                    p_name = str(row.get("page") or "").strip()
                    if p_name:
                        ad_ov[p_name] = {
                            "page": p_name,
                            "daily_budget": _parse_float(row.get("daily_budget")),
                            "boost_pct": str(row.get("boost_pct") or "").strip(),
                            "adjusted_budget": _parse_float(row.get("adjusted_budget")),
                            "boost_date": str(row.get("boost_date") or "").strip(),
                            "monthly_total": _parse_float(row.get("monthly_total")),
                        }
                _save_store(store)
            self._send_json({"ok": True, "state": build_full_state()})
            return

        if parsed.path == "/api/auth/login":
            uname = str(body.get("username") or "").strip().lower()
            pwd = str(body.get("password") or "").strip()
            store = _load_store()
            users = store.get("users") or DEFAULT_USERS
            matched = next(
                (
                    u
                    for u in users
                    if str(u.get("username") or "").strip().lower() == uname
                    and str(u.get("password") or "").strip() == pwd
                ),
                None,
            )
            if not matched:
                self._send_json({"ok": False, "error": "ชื่อผู้ใช้ หรือ รหัสผ่านไม่ถูกต้อง"}, status=401)
                return
            self._send_json({"ok": True, "user": matched})
            return

        if parsed.path == "/api/users/save":
            users_payload = body.get("users") or []
            cleaned: list[dict[str, Any]] = []
            for u in users_payload:
                uname = str(u.get("username") or "").strip()
                if not uname:
                    continue
                cleaned.append({
                    "id": str(u.get("id") or ("usr-" + uuid.uuid4().hex[:8])),
                    "username": uname,
                    "password": str(u.get("password") or "1234").strip() or "1234",
                    "display_name": str(u.get("display_name") or uname).strip(),
                    "tabs": list(u.get("tabs") or ["stock"]),
                    "can_edit": bool(u.get("can_edit", True)),
                    "hide_money": bool(u.get("hide_money", False)),
                    "can_manage_users": bool(u.get("can_manage_users", False)),
                })
            if not any(x.get("can_manage_users") for x in cleaned):
                if cleaned:
                    cleaned[0]["can_manage_users"] = True
                    cleaned[0]["tabs"] = ["overview", "income", "expense", "stock", "customer", "adminboard"]
                else:
                    cleaned = list(DEFAULT_USERS)
            with _LOCK:
                store = _load_store()
                store["users"] = cleaned
                _save_store(store)
            self._send_json({"ok": True, "state": build_full_state()})
            return

        self._send_json({"error": "Unknown endpoint"}, status=404)


_PUBLIC_TUNNEL_URL = ""
PERMANENT_GATEWAY_URL = "https://theearththailand3-crypto.github.io/erp/"


def _get_lan_ip() -> str:
    try:
        s = socket.socket(socket.AF_INET, socket.SOCK_DGRAM)
        s.connect(("8.8.8.8", 80))
        ip = s.getsockname()[0]
        s.close()
        return ip
    except Exception:
        return "127.0.0.1"


def _get_github_token() -> str:
    tok = os.environ.get("GITHUB_TOKEN", "").strip()
    if tok:
        return tok
    try:
        import win32cred

        cred = win32cred.CredRead("git:https://github.com", win32cred.CRED_TYPE_GENERIC)
        return cred["CredentialBlob"].decode("utf-16le").strip()
    except Exception:
        pass
    try:
        import ctypes
        from ctypes import wintypes

        class _CREDENTIALW(ctypes.Structure):
            _fields_ = [
                ("Flags", wintypes.DWORD),
                ("Type", wintypes.DWORD),
                ("TargetName", wintypes.LPWSTR),
                ("Comment", wintypes.LPWSTR),
                ("LastWritten", wintypes.FILETIME),
                ("CredentialBlobSize", wintypes.DWORD),
                ("CredentialBlob", ctypes.POINTER(ctypes.c_ubyte)),
                ("Persist", wintypes.DWORD),
                ("AttributeCount", wintypes.DWORD),
                ("Attributes", ctypes.c_void_p),
                ("TargetAlias", wintypes.LPWSTR),
                ("UserName", wintypes.LPWSTR),
            ]

        pcred = ctypes.POINTER(_CREDENTIALW)()
        if ctypes.windll.advapi32.CredReadW("git:https://github.com", 1, 0, ctypes.byref(pcred)):
            try:
                c = pcred.contents
                raw = ctypes.string_at(c.CredentialBlob, c.CredentialBlobSize)
                return raw.decode("utf-16le").strip()
            finally:
                ctypes.windll.advapi32.CredFree(pcred)
    except Exception:
        pass
    return ""


def _github_put_file(repo_path: str, content_bytes: bytes, msg: str) -> None:
    import base64

    tok = _get_github_token()
    if not tok:
        return
    headers = {
        "Authorization": f"Bearer {tok}",
        "User-Agent": "stprint-erp",
        "Accept": "application/vnd.github+json",
        "Content-Type": "application/json",
    }
    api_url = f"https://api.github.com/repos/theearththailand3-crypto/erp/contents/{repo_path}"
    sha = None
    try:
        r0 = urllib.request.Request(api_url, headers=headers)
        with urllib.request.urlopen(r0, timeout=12) as resp:
            sha = json.loads(resp.read().decode("utf-8")).get("sha")
    except Exception:
        pass
    body: dict[str, Any] = {
        "message": msg,
        "content": base64.b64encode(content_bytes).decode("ascii"),
    }
    if sha:
        body["sha"] = sha
    req = urllib.request.Request(api_url, data=json.dumps(body).encode("utf-8"), headers=headers, method="PUT")
    with urllib.request.urlopen(req, timeout=25):
        pass


def build_compact_cloud_state_b64() -> str:
    import base64
    import gzip

    st = build_full_state()
    for c in st.get("customers") or []:
        c.pop("orders", None)
    raw = json.dumps(st, ensure_ascii=False, separators=(",", ":")).encode("utf-8")
    gz = gzip.compress(raw)
    return base64.b64encode(gz).decode("ascii")


_CLOUD_SYNC_LOCK = threading.Lock()


def _publish_cloud_state() -> None:
    if not _CLOUD_SYNC_LOCK.acquire(blocking=False):
        return
    try:
        b64_str = build_compact_cloud_state_b64()
        (DATA_DIR / "cloud_state.b64").write_text(b64_str, encoding="ascii")
        _github_put_file("cloud_state.b64", b64_str.encode("ascii"), "Sync cloud_state.b64")
        if STORE_PATH.exists():
            _github_put_file("data/store.json", STORE_PATH.read_bytes(), "Sync data/store.json")
    except Exception:
        pass
    finally:
        _CLOUD_SYNC_LOCK.release()


def _publish_cloud_state_async() -> None:
    threading.Thread(target=_publish_cloud_state, daemon=True).start()


def _publish_permanent_gateway(tunnel_url: str) -> None:
    if not tunnel_url:
        return
    try:
        payload = {
            "url": tunnel_url,
            "lan": f"http://{_get_lan_ip()}:{PORT}",
            "updated_at": dt.datetime.now(BKK).isoformat(),
        }
        _github_put_file(
            "live.json",
            json.dumps(payload, indent=2).encode("utf-8"),
            f"Auto-update live ERP tunnel: {tunnel_url}",
        )
    except Exception:
        pass


def _write_mobile_url_file(tunnel_url: str) -> None:
    try:
        (DATA_DIR / "mobile_url.txt").write_text(
            f"ลิงก์ออนไลน์ถาวร (เซฟเก็บไว้ใช้ได้ตลอดชีพ แม้รีสตาร์ทคอม):\n{PERMANENT_GATEWAY_URL}\n\n"
            f"ลิงก์ตรงเซิร์ฟเวอร์รอบนี้ (4G/5G และ Wi-Fi):\n{tunnel_url}\n\n"
            f"ลิงก์ Wi-Fi วงเดียวกัน:\nhttp://{_get_lan_ip()}:{PORT}\n",
            encoding="utf-8",
        )
    except Exception:
        pass


class _ExclusiveHTTPServer(ThreadingHTTPServer):
    allow_reuse_address = False


def _start_cloudflare_tunnel() -> None:
    global _PUBLIC_TUNNEL_URL
    cf_bin = shutil.which("cloudflared")
    if not cf_bin:
        return
    creationflags = 0x08000000 if os.name == "nt" else 0
    sinfo = None
    if os.name == "nt":
        sinfo = subprocess.STARTUPINFO()
        sinfo.dwFlags |= subprocess.STARTF_USESHOWWINDOW
        sinfo.wShowWindow = 0
    while True:
        try:
            proc = subprocess.Popen(
                [cf_bin, "tunnel", "--url", f"http://127.0.0.1:{PORT}"],
                stdout=subprocess.PIPE,
                stderr=subprocess.STDOUT,
                stdin=subprocess.DEVNULL,
                text=True,
                encoding="utf-8",
                errors="replace",
                creationflags=creationflags,
                startupinfo=sinfo,
            )
            if proc.stdout:
                for line in proc.stdout:
                    m = re.search(r"https://[a-zA-Z0-9\-]+\.trycloudflare\.com", line)
                    if m:
                        _PUBLIC_TUNNEL_URL = m.group(0)
                        _write_mobile_url_file(_PUBLIC_TUNNEL_URL)
                        _publish_permanent_gateway(_PUBLIC_TUNNEL_URL)
            proc.wait()
        except Exception:
            pass
        time.sleep(10)


def _auto_sync_loop() -> None:
    try:
        sync_from_google_sheet(all_months=False)
    except Exception:
        pass
    try:
        sync_admin_page_sheet()
    except Exception:
        pass
    try:
        sync_myorder_shipping()
    except Exception:
        pass
    try:
        _publish_cloud_state()
    except Exception:
        pass
    while True:
        time.sleep(180)
        try:
            sync_from_google_sheet(all_months=False)
        except Exception:
            pass
        try:
            sync_admin_page_sheet()
        except Exception:
            pass
        try:
            sync_myorder_shipping()
        except Exception:
            pass
        try:
            _publish_cloud_state()
        except Exception:
            pass


def main() -> None:
    if "--cloud-build" in sys.argv:
        try:
            sync_from_google_sheet(all_months=False)
        except Exception as e:
            print("sheet sync err:", e)
        try:
            sync_admin_page_sheet()
        except Exception as e:
            print("admin sheet sync err:", e)
        try:
            sync_myorder_shipping()
        except Exception as e:
            print("myorder sync err:", e)
        b64_str = build_compact_cloud_state_b64()
        (BASE_DIR / "cloud_state.b64").write_text(b64_str, encoding="ascii")
        print("Built cloud_state.b64:", len(b64_str), "bytes")
        return
    if sys.stdout is None or sys.stderr is None:
        log_fp = open(DATA_DIR / "server.log", "a", encoding="utf-8", errors="replace")
        if sys.stdout is None:
            sys.stdout = log_fp
        if sys.stderr is None:
            sys.stderr = log_fp
    else:
        for stream in (sys.stdout, sys.stderr):
            if hasattr(stream, "reconfigure"):
                try:
                    stream.reconfigure(encoding="utf-8", errors="replace")
                except Exception:
                    pass
    print("========================================================")
    print("ST Print - ระบบเช็คข้อมูลร้าน (รายรับ / รายจ่าย / สต๊อก / ลูกค้า)")
    print(f"เปิดหน้าเว็บในคอมที่: http://localhost:{PORT}")
    print(f"เปิดหน้าเว็บในมือถือ (Wi-Fi เดียวกัน): http://{_get_lan_ip()}:{PORT}")
    print("========================================================")
    if not SHEET_CACHE_PATH.exists():
        try:
            print("กำลังซิงก์ข้อมูลล่าสุดจาก Google Sheet ครั้งแรก...")
            sync_from_google_sheet(all_months=False)
            print("ซิงก์ข้อมูล Google Sheet เรียบร้อย!")
        except Exception as e:
            print(f"ซิงก์ครั้งแรกไม่สำเร็จ ({e}) — สามารถกดซิงก์ในหน้าเว็บได้")
    try:
        server = _ExclusiveHTTPServer(("0.0.0.0", PORT), DashboardHandler)
    except OSError:
        print(f"เซิร์ฟเวอร์ทำงานอยู่แล้วที่ http://localhost:{PORT}")
        return
    threading.Thread(target=_auto_sync_loop, daemon=True).start()
    threading.Thread(target=_start_cloudflare_tunnel, daemon=True).start()
    try:
        server.serve_forever()
    except KeyboardInterrupt:
        pass


if __name__ == "__main__":
    main()
