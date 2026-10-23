# -*- coding: utf-8 -*-
"""
محرّك التدقيق المالي الذكي — Intelligent Financial Audit Engine
================================================================
قراءة مرنة للملفات الحكومية (Excel / XLSB / XLS / CSV) متعددة الأوراق،
تصنيف الأوراق تلقائياً (مواقف مالية / سجلات حركات)، توحيد الأعمدة مهما
اختلفت أسماؤها أو ترتيبها أو لغتها (عربي/إنجليزي)، ثم تطبيق قواعد تدقيق
مبنية على المخاطر وفق المعايير الدولية (ISA 240/315، ISSAI 1240/1260).

الفكرة الأساسية: لا نعتمد على أسماء أعمدة ثابتة، بل نبني لكل عمود نصاً
ترويسياً مجمّعاً ثم نطابقه مع قاموس مرادفات (Aliases) عربي/إنجليزي لتحديد
معناه الوظيفي (تاريخ، مستفيد، مبلغ، مخصصات، أوامر مالية، نسبة إنفاق...).
"""

from __future__ import annotations

import io
import os
import re
import math
import datetime as _dt
from typing import Any

import pandas as pd

# ---------------------------------------------------------------------------
# ثوابت عامة
# ---------------------------------------------------------------------------
SEVERITY_AR = {"high": "عالية", "medium": "متوسطة", "low": "منخفضة"}
SEVERITY_EN = {"high": "High", "medium": "Medium", "low": "Low"}
SEV_WEIGHT = {"high": 7.0, "medium": 3.2, "low": 1.1}

# سقف عدد العناصر المحفوظة داخل كل ملاحظة (لتقليل حجم الاستجابة)
MAX_ITEMS_PER_FINDING = 60

# ---------------------------------------------------------------------------
# أدوات معالجة النصوص والأرقام
# ---------------------------------------------------------------------------
_AR_DIGITS = str.maketrans("٠١٢٣٤٥٦٧٨٩۰۱۲۳۴۵۶۷۸۹", "01234567890123456789")


def norm_digits(s: str) -> str:
    """توحيد الأرقام العربية/الفارسية إلى أرقام لاتينية."""
    return (s or "").translate(_AR_DIGITS)


def compact(s: Any) -> str:
    """نسخة مضغوطة من النص للمطابقة: بلا حركات/تطويل/مسافات/ترقيم."""
    s = norm_digits("" if s is None else str(s))
    s = re.sub(r"[\u064B-\u0652\u0640]", "", s)  # حركات + تطويل
    s = (
        s.replace("أ", "ا")
        .replace("إ", "ا")
        .replace("آ", "ا")
        .replace("ٱ", "ا")
        .replace("ى", "ي")
        .replace("ئ", "ي")
        .replace("ؤ", "و")
        .replace("ة", "ه")
    )
    s = re.sub(r"[^\w\u0600-\u06FF]+", "", s, flags=re.UNICODE)
    return s.lower()


def clean_text(v: Any) -> str:
    """نص نظيف للعرض (بلا قيم NaN)."""
    if v is None:
        return ""
    try:
        if pd.isna(v):
            return ""
    except Exception:
        pass
    return re.sub(r"\s+", " ", str(v)).strip()


def to_number(v: Any) -> float | None:
    """تحويل قيمة خلية إلى رقم آمن (يدعم الفواصل والأقواس والأرقام العربية)."""
    if v is None:
        return None
    if isinstance(v, bool):
        return None
    if isinstance(v, (int, float)):
        try:
            if pd.isna(v):
                return None
        except Exception:
            pass
        return float(v)
    s = norm_digits(str(v)).strip()
    if not s:
        return None
    neg = False
    if s.startswith("(") and s.endswith(")"):
        neg, s = True, s[1:-1]
    s = s.replace(",", "").replace("٬", "").replace(" ", "").replace("\u00a0", "")
    for token in ("%", "د.أ", "JOD", "jod", "$", "USD", "دينار", "ألف"):
        s = s.replace(token, "")
    if not re.fullmatch(r"[+-]?\d+(\.\d+)?", s):
        # نحاول استخراج أول رقم صالح
        m = re.search(r"[+-]?\d+(?:\.\d+)?", s)
        if not m:
            return None
        s = m.group(0)
    try:
        val = float(s)
    except Exception:
        return None
    return -val if neg else val


def to_date(v: Any) -> _dt.date | None:
    """تحويل قيمة خلية إلى تاريخ (يدعم تواريخ Excel الرقمية والنصوص)."""
    if v is None:
        return None
    if isinstance(v, _dt.datetime):
        return v.date()
    if isinstance(v, _dt.date):
        return v
    if isinstance(v, (int, float)) and not isinstance(v, bool):
        try:
            if pd.isna(v):
                return None
            base = _dt.datetime(1899, 12, 30)
            n = float(v)
            if 20000 < n < 60000:  # نطاق تواريخ Excel المعقولة
                return (base + _dt.timedelta(days=n)).date()
        except Exception:
            pass
        return None
    s = norm_digits(str(v)).strip()
    if not s:
        return None
    s_norm = s.replace("T", " ").split(" ")[0]
    s_norm = s_norm.replace(".", "/").replace("-", "/")
    for fmt in ("%Y/%m/%d", "%d/%m/%Y", "%m/%d/%Y", "%Y/%m", "%Y%m%d", "%d/%m/%y"):
        try:
            return _dt.datetime.strptime(s_norm, fmt).date()
        except Exception:
            continue
    try:
        parsed = pd.to_datetime(s, errors="coerce", dayfirst=True)
        if pd.notna(parsed):
            return parsed.date()
    except Exception:
        pass
    return None


# ---------------------------------------------------------------------------
# قاموس المرادفات (عربي + إنجليزي) — الأطول مطابقةً هو الأقوى
# ---------------------------------------------------------------------------
LEDGER_ALIASES: dict[str, list[str]] = {
    "date": ["تاريخ الشيك", "تاريخ الدفع", "تاريخ الفاتورة", "التاريخ", "تاريخ", "date"],
    "beneficiary": [
        "اسم الدافع/ المستفيد", "اسم الدافع", "اسم المستفيد", "رقم و إسم المستفيد",
        "المستفيــد", "المستفيد", "المستفيـــد", "واسم المستفيد", "الجهة", "المورد",
        "beneficiary", "vendor_id", "vendor", "payee", "supplier",
    ],
    "amount": ["مبلغ الفاتورة", "مبلغ السطر", "قيمة الشيك", "المبلغ", "القيمة", "amount", "value"],
    "gross": ["اجمالي النفقات", "إجمالي النفقات", "اجمالي", "gross", "total"],
    "deductions": ["اجمالي الاقتطاعات", "الاقتطاعات", "اقتطاعات", "deductions", "withholding"],
    "advances": ["اجمالي سداد السلف", "سداد السلف", "السلف", "advances"],
    "net_paid": ["المبلغ المدفوع", "قيمة الشيك", "المدفوع", "net", "paid_amount"],
    "check_no": ["رقم الشيك", "الشيك", "check", "cheque"],
    "entry_no": ["رقم القيد", "القيد", "entry", "journal"],
    "auto_no": ["الرقم الالي/رقم الفاتورة", "الرقم الالي", "الرقم الآلي", "الالي", "auto", "serial"],
    "spend_no": ["رقم الصرف", "الصرف", "spend", "payment_no"],
    "receipt_no": ["أمر القبض", "امر القبض", "الوصول", "receipt"],
    "status": ["حالة المستند", "حالة", "status"],
    "posting": ["حالة الترحيل", "الترحيل", "posting"],
    "source": ["المصــدر", "المصدر", "source"],
    "description": ["التفاصيــل", "البيــــــــــــان", "البيان", "الوصف", "details", "description"],
    "invoice_no": ["رقم الفاتورة", "الفاتورة", "invoice_number", "invoice"],
    "invoice_type": ["نوع الفاتورة", "invoice_type"],
    "debit": ["المبلغ مدين", "مديـــن", "مدين", "debit"],
    "credit": ["المبلغ دائن", "دائــــن", "دائن", "credit"],
    "account_dist": ["الحساب/المبلغ", "توزيع الحساب", "account"],
    # حقول متوافقة مع الملفات العامة القديمة
    "transaction_id": ["transaction_id", "رقم الحركة"],
}

POSITION_ALIASES: dict[str, list[str]] = {
    "account_code": ["الرقم – الوصف", "الرقم - الوصف", "الرمز الحسابي", "رقم البند", "رقم الحساب", "account_code", "code"],
    "item_name": ["الرقم – الوصف", "الرقم - الوصف", "الوصف", "الحساب", "البيان", "item", "description"],
    "obligations": ["التزامات الشراء", "إلتزامات", "التزامات", "obligations", "commitment"],
    "custody": ["حوالات نقل العهدة", "نقل العهدة", "custody"],
    "estimated_budget": ["المخصصات المقدر", "المخصصات", "المقدر", "estimated", "appropriation", "budget"],
    "financial_orders": ["الأوامر المالية", "الاوامر المالية", "الأوامر", "orders"],
    "adjusted_estimate": ["المقدر وتعديلاته", "المقدر وتعديل", "adjusted"],
    "total_transfers": ["الإجمالي لتاريخه الحوالات المالية", "المعاملات الجارية خلال الشهر الحوالات المالية", "الحوالات المالية", "transfers"],
    "total_paid": ["الإجمالي لتاريخه المدفوع", "المعاملات الجارية خلال الشهر المدفوع", "المدفوع", "paid"],
    "total_unpaid": ["غير المدفوع", "unpaid"],
    "total_presented": ["المقدم للصرف", "المقدم للصرف", "presented"],
    "balance_licensed": ["رصيد المخصصات المرخص", "( حوالات)", "المرخص به", "licensed_balance"],
    "balance_free": ["رصيد المخصصات الحر", "الرصيد الحر", "balance", "available"],
    "spend_ratio": ["نسبة انفاق", "نسبة الإنفاق", "نسبه الانفاق", "spend_ratio", "utilization", "ratio"],
}

# كلمات تُشير إلى أن الصف ترويسي/إجمالي/وصفي وليس بيانات
TOTAL_WORDS = ["الاجمالي", "المجموع", "مجموع", "total", "رصيد اول المده", "الرصيد المدور", "الصفحة", "المستخدم"]


# ---------------------------------------------------------------------------
# قراءة الملفات (متعددة الأوراق) — Excel / XLSB / XLS / CSV
# ---------------------------------------------------------------------------
def _read_csv_bytes(contents: bytes) -> dict[str, pd.DataFrame]:
    last_err: Exception | None = None
    for enc in ("utf-8-sig", "utf-8", "cp1256", "latin-1"):
        for sep in (None, ",", ";", "\t", "|"):
            try:
                df = pd.read_csv(
                    io.BytesIO(contents), header=None, dtype=object,
                    encoding=enc, sep=sep, engine="python",
                )
                return {"CSV": df}
            except Exception as exc:  # noqa: BLE001
                last_err = exc
    raise ValueError(f"csv_read_failed: {last_err}")


def read_workbook(filename: str, contents: bytes) -> dict[str, pd.DataFrame]:
    """يُعيد قاموساً: اسم الورقة -> DataFrame خام (header=None)."""
    name = (filename or "").lower()
    if name.endswith(".csv") or name.endswith(".txt"):
        return _read_csv_bytes(contents)
    if name.endswith((".xlsx", ".xlsm", ".xls", ".xlsb")):
        buf = io.BytesIO(contents)
        try:
            xls = pd.ExcelFile(buf)
        except Exception as exc:  # noqa: BLE001
            raise ValueError(f"excel_read_failed: {exc}")
        out: dict[str, pd.DataFrame] = {}
        for sn in xls.sheet_names:
            try:
                buf.seek(0)
                out[sn] = pd.read_excel(xls, sheet_name=sn, header=None, dtype=object)
            except Exception:  # noqa: BLE001
                continue
        if not out:
            raise ValueError("excel_read_failed: no_readable_sheets")
        return out
    raise ValueError("unsupported_format")


# ---------------------------------------------------------------------------
# كشف الترويسة وتوحيد الأعمدة
# ---------------------------------------------------------------------------
def _row_token_hits(row_values: list[Any], alias_map: dict[str, list[str]]) -> int:
    """عدد الأعمدة في الصف التي تطابق أي مرادف (لاختيار صف الترويسة)."""
    hits = 0
    for cell in row_values:
        h = compact(cell)
        if not h or len(h) > 60:
            continue
        if any(compact(a) in h for aliases in alias_map.values() for a in aliases):
            hits += 1
    return hits


def _best_alias_match(header_compact: str, aliases: list[str]) -> int:
    best = 0
    for alias in aliases:
        a = compact(alias)
        if a and a in header_compact:
            best = max(best, len(a))
    return best


def detect_columns(
    raw: pd.DataFrame, alias_map: dict[str, list[str]]
) -> tuple[dict[int, str], int, int]:
    """
    يحدد لكل عمود الحقل الوظيفي، ويعيد (خريطة الأعمدة، صف الترويسة، صف بدء البيانات).
    """
    n = len(raw)
    if n == 0:
        return {}, 0, 0
    scan = min(n, 40)
    hits = [(_row_token_hits(list(raw.iloc[r].values), alias_map), r) for r in range(scan)]
    best_hits, hdr = max(hits, key=lambda t: t[0])
    if best_hits == 0:
        hdr = 0

    # أول صف بعد الترويسة يحتوي على قيمة رقمية => بداية البيانات
    data_start = n
    for r in range(hdr + 1, min(n, hdr + 10)):
        vals = list(raw.iloc[r].values)
        if any(to_number(v) is not None for v in vals):
            data_start = r
            break

    header_end = max(hdr, data_start - 1 if data_start < n else hdr)
    lo = max(0, hdr - 2)
    col_text: dict[int, str] = {}
    for c in range(raw.shape[1]):
        parts: list[str] = []
        for r in range(lo, min(header_end + 1, n)):
            t = clean_text(raw.iat[r, c])
            if t:
                parts.append(t)
        col_text[c] = " ".join(parts)

    mapping: dict[int, str] = {}
    for c, text in col_text.items():
        hc = compact(text)
        if not hc:
            continue
        best_field, best_score = None, 0
        for field, aliases in alias_map.items():
            sc = _best_alias_match(hc, aliases)
            if sc > best_score:
                best_field, best_score = field, sc
        if best_field:
            mapping[c] = best_field
    return mapping, hdr, data_start


def _sheet_kind(pos_map: dict[int, str], led_map: dict[int, str]) -> str:
    """يقرر نوع الورقة (موقف مالي / سجل حركات) بمقارنة قوة مطابقة الأعمدة."""
    pos_fields = set(pos_map.values())
    led_fields = set(led_map.values())
    pos_keys = {"estimated_budget", "financial_orders", "spend_ratio", "balance_free", "adjusted_estimate", "total_paid"}
    led_keys = {"date", "beneficiary", "amount", "debit", "credit", "check_no", "net_paid"}

    p = len(pos_fields & pos_keys)
    l = len(led_fields & led_keys)
    # مؤشرات حاسمة لنوع الورقة
    if "spend_ratio" in pos_fields:
        p += 3
    if "estimated_budget" in pos_fields and "financial_orders" in pos_fields:
        p += 2
    if "debit" in led_fields or "credit" in led_fields:
        l += 2
    if "beneficiary" in led_fields and "date" in led_fields:
        l += 1

    if p == 0 and l == 0:
        return "unknown"
    return "position" if p >= l else "ledger"


def _first_col(mapping: dict[int, str], field: str) -> int | None:
    for c, f in mapping.items():
        if f == field:
            return c
    return None


# ---------------------------------------------------------------------------
# استخراج البنود (المواقف المالية) وسجلات الحركات
# ---------------------------------------------------------------------------
def extract_position_items(raw: pd.DataFrame, mapping: dict[int, str], data_start: int, sheet: str) -> list[dict]:
    code_col = _first_col(mapping, "account_code") or _first_col(mapping, "item_name")
    name_col = _first_col(mapping, "item_name") or code_col
    num_cols = {f: c for c, f in mapping.items() if f in POSITION_ALIASES and f not in ("account_code", "item_name")}
    year = None
    m = re.search(r"(20\d{2})", str(sheet))
    if m:
        year = int(m.group(1))

    items: list[dict] = []
    current: dict | None = None
    for r in range(data_start, len(raw)):
        row = raw.iloc[r]
        code_raw = row[code_col] if code_col is not None else None
        name_raw = row[name_col] if name_col is not None else None
        code_num = to_number(code_raw)
        code_txt = clean_text(code_raw)
        name_txt = clean_text(name_raw)

        digits = code_txt.replace(".", "").replace(",", "")
        if isinstance(code_raw, (int, float)) and not isinstance(code_raw, bool):
            is_code_row = code_num is not None
        else:
            # صف الرمز الحسابي: نص يتكون من أرقام فقط (رمز حساب طويل)
            is_code_row = bool(digits) and digits.isdigit() and len(digits) >= 4

        if is_code_row:
            current = {"sheet": sheet, "year": year, "account_code": code_txt}
            for field, c in num_cols.items():
                current[field] = to_number(row[c])
            items.append(current)
            continue

        # صف الوصف (يلي صف الرمز) أو صف نصي آخر
        if name_txt and not any(w in compact(name_txt) for w in [compact(x) for x in TOTAL_WORDS]):
            if current is not None and not current.get("item_name"):
                current["item_name"] = name_txt
            else:
                # صف وصفي مستقل بلا رمز
                items.append({"sheet": sheet, "year": year, "account_code": "", "item_name": name_txt})

    # تنظيف وتصفية
    out = []
    for it in items:
        if not it.get("item_name") and not it.get("account_code"):
            continue
        est = it.get("estimated_budget")
        paid = it.get("total_paid")
        ratio = it.get("spend_ratio")
        if ratio is None and est and est > 0 and paid is not None:
            ratio = round(paid / est * 100, 1)
        it["spend_ratio"] = ratio
        out.append(it)
    return out


def extract_ledger_records(raw: pd.DataFrame, mapping: dict[int, str], data_start: int, sheet: str) -> list[dict]:
    field_col = {f: c for c, f in mapping.items()}
    recs: list[dict] = []

    def get(row, field):
        c = field_col.get(field)
        return row[c] if c is not None else None

    for r in range(data_start, len(raw)):
        row = raw.iloc[r]
        d = to_date(get(row, "date"))
        debit = to_number(get(row, "debit"))
        credit = to_number(get(row, "credit"))
        amount = to_number(get(row, "amount"))
        net = to_number(get(row, "net_paid"))
        gross = to_number(get(row, "gross"))
        deductions = to_number(get(row, "deductions"))
        adv = to_number(get(row, "advances"))

        # المبلغ الفعلي للحركة
        eff = amount
        if eff is None:
            eff = net
        if eff is None:
            eff = gross
        if eff is None:
            eff = debit if debit is not None else credit

        beneficiary = clean_text(get(row, "beneficiary"))
        has_identity = (
            (d is not None and (beneficiary or eff is not None))
            or (eff is not None and beneficiary)
        )
        # ضم الصفوف التكميلية (رقم الشيك، توزيع الحسابات) للسجل السابق
        if not has_identity and recs:
            prev = recs[-1]
            for field in ("check_no", "entry_no", "spend_no", "receipt_no", "auto_no", "posting", "status", "source"):
                c = field_col.get(field)
                if c is None:
                    continue
                val = clean_text(row[c])
                if val and not prev.get(field):
                    prev[field] = val
            c = field_col.get("account_dist")
            if c is not None:
                val = clean_text(row[c])
                if val:
                    prev.setdefault("account_dist", [])
                    if isinstance(prev["account_dist"], list):
                        prev["account_dist"].append(val)
            continue

        rec: dict[str, Any] = {
            "sheet": sheet,
            "date": d.isoformat() if d else "",
            "beneficiary": beneficiary,
            "amount": eff,
            "gross": gross,
            "net": net,
            "deductions": deductions,
            "advances": adv,
            "debit": debit,
            "credit": credit,
            "check_no": clean_text(get(row, "check_no")),
            "entry_no": clean_text(get(row, "entry_no")),
            "spend_no": clean_text(get(row, "spend_no")),
            "receipt_no": clean_text(get(row, "receipt_no")),
            "auto_no": clean_text(get(row, "auto_no")),
            "status": clean_text(get(row, "status")),
            "posting": clean_text(get(row, "posting")),
            "source": clean_text(get(row, "source")),
            "description": clean_text(get(row, "description")),
            "invoice_no": clean_text(get(row, "invoice_no")),
            "transaction_id": clean_text(get(row, "transaction_id")),
        }
        recs.append(rec)
    return recs


def _norm_beneficiary(name: str) -> str:
    n = compact(name)
    # إزالة الأرقام والرموز البادئة (مثل 815900001\)
    n = re.sub(r"^[0-9\\/-]+", "", n)
    return n


# ---------------------------------------------------------------------------
# بناء الملاحظات (Findings)
# ---------------------------------------------------------------------------
def _finding(fid, severity, title_ar, title_en, desc_ar, desc_en, standard, items, impact=None) -> dict:
    return {
        "id": fid,
        "rule": fid,
        "severity": severity,
        "severity_ar": SEVERITY_AR[severity],
        "severity_en": SEVERITY_EN[severity],
        "title_ar": title_ar,
        "title_en": title_en,
        "desc_ar": desc_ar,
        "desc_en": desc_en,
        "standard": standard,
        "count": len(items),
        "impact": impact,
        "items": items[:MAX_ITEMS_PER_FINDING],
    }


def _audit_position(items: list[dict], metrics: dict) -> list[dict]:
    findings: list[dict] = []

    overspend, near, idle, over_orders, neg_balance, low_exec, amended, zero_budget_paid = [], [], [], [], [], [], [], []

    for it in items:
        est = it.get("estimated_budget") or 0.0
        orders = it.get("financial_orders") or 0.0
        adjusted = it.get("adjusted_estimate") or 0.0
        paid = it.get("total_paid") or 0.0
        bal_free = it.get("balance_free")
        ratio = it.get("spend_ratio")
        row = {
            "sheet": it.get("sheet"),
            "year": it.get("year"),
            "account_code": it.get("account_code"),
            "item_name": it.get("item_name"),
            "estimated_budget": est,
            "financial_orders": orders,
            "paid": paid,
            "balance": bal_free,
            "spend_ratio": ratio,
        }
        if ratio is not None and ratio > 100.5:
            overspend.append(row)
        elif ratio is not None and ratio >= 90:
            near.append(row)
        if est > 0 and orders == 0 and paid == 0:
            idle.append(row)
        if orders and paid > orders + 0.5:
            over_orders.append(row)
        if bal_free is not None and bal_free < -0.5:
            neg_balance.append(row)
        if est >= 20000 and ratio is not None and 0 < ratio < 30:
            low_exec.append(row)
        if est > 0 and adjusted and abs(adjusted - est) > max(1.0, est * 0.001):
            amended.append({**row, "adjusted": adjusted, "delta": round(adjusted - est, 2)})
        if est == 0 and paid > 0:
            zero_budget_paid.append(row)

    if overspend:
        findings.append(_finding(
            "POS_OVERSPEND", "high",
            "بنود تجاوزت المخصصات المقدرة (نسبة إنفاق ≥ 100%)",
            "Line items exceeding the approved appropriation (spend ratio ≥ 100%)",
            "صرف فعلي يفوق المخصصات المقدرة/الأوامر المالية، وهو مؤشر على تجاوز الصلاحيات المالية ومخالفة قواعد الانضباط المالي.",
            "Actual expenditure exceeds the approved appropriation/authorization, indicating potential breach of financial authority.",
            "ISA 240 / ISSAI 1240 (Compliance & regularity)",
            overspend,
            impact=round(sum(r["paid"] - r["estimated_budget"] for r in overspend), 2),
        ))
    if neg_balance:
        findings.append(_finding(
            "POS_NEG_BALANCE", "high",
            "أرصدة سالبة في المخصصات الحرة",
            "Negative free-appropriation balances",
            "ظهور رصيد سالب يعني صرفاً بمبالغ أكبر من المتاح، ما قد يشير إلى خطأ محاسبي أو صرف غير مبرر.",
            "A negative available balance indicates over-commitment of available funds.",
            "ISSAI 1260 / internal control",
            neg_balance,
        ))
    if zero_budget_paid:
        findings.append(_finding(
            "POS_ZERO_BUDGET_PAID", "medium",
            "صرف على بنود بمخصصات مقدرة = صفر",
            "Expenditure on items with zero estimated appropriation",
            "وجود مدفوعات فعلية على بنود مخصصاتها المقدرة صفر (تمويل بالأوامر المالية فقط) يستوجب توثيق مصدر التمويل وصلاحيته.",
            "Actual payments recorded against a zero estimated appropriation — the funding source and authority should be documented.",
            "ISA 315 / ISSAI 1240",
            zero_budget_paid,
        ))
    if over_orders:
        findings.append(_finding(
            "POS_OVER_ORDERS", "high",
            "مدفوعات تتجاوز الأوامر المالية",
            "Payments exceeding the financial orders",
            "الصرف الفعلي أكبر من قيمة الأوامر المالية المخصصة، ما يستوجب التحقق من مبررات التجاوز.",
            "Payments exceed the value of issued financial orders and require justification.",
            "ISSAI 1240",
            over_orders,
        ))
    if near:
        findings.append(_finding(
            "POS_NEAR_CEILING", "medium",
            "بنود قاربت سقف الإنفاق (90% – 100%)",
            "Items approaching the spending ceiling (90%–100%)",
            "استهلاك شبه كامل للمخصصات مع مخاطر ضغط مالي في الفترات المتبقية من السنة المالية.",
            "Near-full appropriation utilization creates budget pressure risk.",
            "ISA 315 (risk of material misstatement) / ISSAI 1240",
            near,
        ))
    if idle:
        findings.append(_finding(
            "POS_IDLE_ALLOCATION", "medium",
            "مخصصات دون أي حركة مالية (بنود معطلة)",
            "Allocations with no financial movement (idle items)",
            "تم تخصيص مخصصات لكن لم تُصدر أوامر مالية أو مدفوعات، ما يشير إلى ضعف التخطيط أو تعطّل التنفيذ.",
            "Allocations received but no orders/payments were issued — weak planning or stalled execution.",
            "ISSAI 1260 (Economy & efficiency)",
            idle,
        ))
    if amended:
        findings.append(_finding(
            "POS_AMENDMENT", "medium",
            "تعديلات على المخصصات المقدرة",
            "Amendments to estimated appropriations",
            "اختلاف قيمة 'المقدر وتعديلاته' عن 'المقدر' الأصلي، ويتطلب تتبع مسوّغات التعديل وصلاحياته.",
            "The adjusted estimate differs from the original estimate and requires authorization trail.",
            "ISA 240 (management override) / ISSAI 1240",
            amended,
        ))
    if low_exec:
        findings.append(_finding(
            "POS_LOW_EXECUTION", "low",
            "انخفاض نسبة التنفيذ (< 30%) مقابل مخصصات مرتفعة",
            "Low execution rate (< 30%) against sizeable allocations",
            "ضعف قدرة الجهة على الإنفاق بما يتوافق مع المخصصات المعتمدة (مخاطر كفاءة/فعالية).",
            "Low budget execution suggests efficiency/effectiveness risk.",
            "ISSAI 1260",
            low_exec,
        ))

    # تحليل الاتجاه السنوي (Year-over-Year)
    yoy = _yoy_findings(items)
    findings.extend(yoy)
    return findings


def _yoy_findings(items: list[dict]) -> list[dict]:
    """مقارنة سنوية بين نسب الإنفاق والمخصصات للبنود المتشابهة."""
    by_code: dict[str, dict[int, dict]] = {}
    for it in items:
        code = it.get("account_code") or compact(it.get("item_name") or "")
        year = it.get("year")
        if not code or not year:
            continue
        by_code.setdefault(code, {})[year] = it

    jumps = []
    for code, years in by_code.items():
        seq = sorted(years.keys())
        for a, b in zip(seq, seq[1:]):
            ia, ib = years[a], years[b]
            ra, rb = ia.get("spend_ratio"), ib.get("spend_ratio")
            ba, bb = ia.get("estimated_budget") or 0, ib.get("estimated_budget") or 0
            if ra is not None and rb is not None and abs(rb - ra) >= 30:
                jumps.append({
                    "account_code": code,
                    "item_name": ib.get("item_name") or ia.get("item_name"),
                    "from_year": a, "to_year": b,
                    "ratio_from": ra, "ratio_to": rb,
                    "delta": round(rb - ra, 1),
                    "budget_from": ba, "budget_to": bb,
                })
    if jumps:
        return [_finding(
            "POS_YOY_VARIANCE", "medium",
            "تباينات سنوية جوهرية في نسب الإنفاق",
            "Significant year-over-year variances in spending ratios",
            "تغيّر نسبة الإنفاق بأكثر من 30 نقطة بين سنة وأخرى على نفس البند، ما يستدعي التحليل والمساءلة.",
            "Spending ratio shifted by more than 30 points year-over-year on the same item.",
            "ISA 315 (analytical procedures)",
            jumps,
        )]
    return []


def _audit_ledger(
    records: list[dict], threshold: float = 10000.0, invoice_identifier: bool = False
) -> tuple[list[dict], dict, dict]:
    findings: list[dict] = []
    positive = [r for r in records if isinstance(r.get("amount"), (int, float)) and float(r["amount"]) > 0]
    total = len(positive)
    total_amount = float(sum(r["amount"] for r in positive)) if positive else 0.0
    avg_amount = round(total_amount / total, 2) if total else 0.0

    # --- 1) مدفوعات مكررة (نفس المستفيد + نفس المبلغ خلال نافذة زمنية) ---
    dups: list[dict] = []
    buckets: dict[tuple, list[dict]] = {}
    for r in positive:
        key = (_norm_beneficiary(r.get("beneficiary") or ""), round(float(r["amount"]), 2))
        buckets.setdefault(key, []).append(r)
    for (ben, amt), group in buckets.items():
        if not ben or len(group) < 2:
            continue
        dates = sorted(d for d in (to_date(g.get("date")) for g in group) if d)
        span_ok = (not dates) or ((dates[-1] - dates[0]).days <= 30)
        if span_ok:
            dups.append({
                "beneficiary": group[0].get("beneficiary"),
                "amount": round(float(amt), 2),
                "occurrences": len(group),
                "dates": [d.isoformat() for d in dates][:10],
                "check_no": [g.get("check_no") for g in group][:10],
                "sheets": sorted({g.get("sheet") for g in group}),
            })
    if dups:
        findings.append(_finding(
            "LED_DUPLICATE_PAYMENTS", "high",
            "مدفوعات مكررة لنفس المستفيد وبالمبلغ ذاته",
            "Duplicate payments to the same beneficiary for the same amount",
            "تكرار صرف بنفس المستفيد ونفس القيمة خلال فترة قصيرة يعتبر مؤشراً رئيسياً على مخاطر الاحتيال (Fraud) أو الخطأ في الصرف.",
            "Repeated identical payments to the same beneficiary indicate a key fraud risk (ISA 240).",
            "ISA 240 (Fraud) — presumed risk of fraudulent disbursements",
            dups,
            impact=round(sum(d["amount"] * (d["occurrences"] - 1) for d in dups), 2),
        ))

    # --- 2) تجزئة المشتريات (Split purchases) ---
    weekly: dict[tuple, float] = {}
    weekly_rows: dict[tuple, list] = {}
    for r in positive:
        d = to_date(r.get("date"))
        if not d:
            continue
        wk = d - _dt.timedelta(days=d.weekday())  # بداية الأسبوع (الإثنين)
        key = (_norm_beneficiary(r.get("beneficiary") or ""), wk)
        weekly[key] = weekly.get(key, 0.0) + float(r["amount"])
        weekly_rows.setdefault(key, []).append(r)
    splits = []
    for (ben, wk), s in weekly.items():
        if ben and s > threshold and len(weekly_rows[(ben, wk)]) > 1:
            splits.append({
                "beneficiary": weekly_rows[(ben, wk)][0].get("beneficiary"),
                "week_start": wk.isoformat(),
                "count": len(weekly_rows[(ben, wk)]),
                "weekly_total": round(s, 2),
                "threshold": threshold,
            })
    splits.sort(key=lambda x: x["weekly_total"], reverse=True)
    if splits:
        findings.append(_finding(
            "LED_SPLIT_PURCHASES", "medium",
            "تجزئة المشتريات لتجاوز سقف الشراء المباشر",
            "Split purchases to bypass the direct-purchase ceiling",
            f"مجموع مدفوعات مستفيد واحد خلال أسبوع واحد تجاوز السقف ({threshold:,.0f}) عبر عدة حركات، وهو نمط شائع لتجاوز حدود الصلاحيات.",
            f"Weekly cumulative payments to one beneficiary exceeded the ceiling ({threshold:,.0f}) via multiple transactions.",
            "ISA 240 (fraud) / procurement control",
            splits,
            impact=round(sum(s["weekly_total"] for s in splits), 2),
        ))

    # --- 3) حركات عطلة نهاية الأسبوع (الجمعة/السبت) ---
    weekend = []
    for r in records:
        d = to_date(r.get("date"))
        if d and d.weekday() in (4, 5):
            weekend.append({
                "date": d.isoformat(),
                "day": "الجمعة" if d.weekday() == 4 else "السبت",
                "beneficiary": r.get("beneficiary"),
                "amount": r.get("amount"),
                "check_no": r.get("check_no"),
                "sheet": r.get("sheet"),
            })
    if weekend:
        findings.append(_finding(
            "LED_WEEKEND_TRANSACTIONS", "medium",
            "حركات مالية في أيام العطلة (الجمعة/السبت)",
            "Transactions on weekends (Friday/Saturday)",
            "تنفيذ حركات مالية في أيام العطلة الرسمية قد يشير إلى معاملات خارج أوقات العمل الاعتيادية أو سعي لتقليل الرقابة.",
            "Financial transactions executed on official weekends may indicate unusual activity.",
            "ISA 240 (unusual transactions)",
            weekend,
        ))

    # --- 4) مبالغ سالبة (قيود عكسية / تجاوزات) ---
    negatives = [r for r in records if isinstance(r.get("amount"), (int, float)) and float(r["amount"]) < 0]
    if negatives:
        findings.append(_finding(
            "LED_NEGATIVE_AMOUNTS", "high",
            "مبالغ سالبة في سجل الصرف",
            "Negative amounts in the payments ledger",
            "وجود مبالغ سالبة يوحي بقيود عكسية أو استردادات غير موثقة، ويجب التحقق من مسوّغاتها.",
            "Negative amounts suggest reversals or undocumented refunds that require verification.",
            "ISA 315 / ISA 240 (data integrity & fraud)",
            [{"date": r.get("date"), "beneficiary": r.get("beneficiary"), "amount": r.get("amount"), "sheet": r.get("sheet")} for r in negatives[:60]],
        ))

    # --- 5) حركات بمبلغ صفر ---
    zeros = [r for r in records if isinstance(r.get("amount"), (int, float)) and float(r["amount"]) == 0]
    if zeros:
        findings.append(_finding(
            "LED_ZERO_AMOUNTS", "low",
            "حركات مالية بمبلغ صفر",
            "Zero-amount transactions",
            "تسجيل حركات بمبلغ صفر قد يعكس إدخالات ناقصة أو حركات إدارية بحاجة إلى توضيح.",
            "Zero-amount entries may reflect incomplete data entry or administrative entries needing clarification.",
            "ISA 315 (data integrity)",
            [{"date": r.get("date"), "beneficiary": r.get("beneficiary"), "check_no": r.get("check_no"), "sheet": r.get("sheet")} for r in zeros[:40]],
        ))

    # --- 6) مبالغ مقرّبة بشكل مثير (Round numbers) ---
    rounds = [r for r in positive if float(r["amount"]) >= 1000 and float(r["amount"]) % 1000 == 0]
    if len(rounds) >= 3:
        findings.append(_finding(
            "LED_ROUND_AMOUNTS", "low",
            "مبالغ مقرّبة بدرجة عالية (من مضاعفات 1000)",
            "Highly rounded amounts (multiples of 1,000)",
            "كثرة المبالغ المقرّبة قد تشير إلى تقديرات غير دقيقة أو صرف بلا مستندات فعلية.",
            "A high frequency of rounded amounts may indicate estimation rather than actual invoicing.",
            "ISSAI 1240 (analytical review)",
            [{"date": r.get("date"), "beneficiary": r.get("beneficiary"), "amount": r.get("amount")} for r in rounds[:40]],
        ))

    # --- 7) مبالغ قاربت سقف الصلاحية — مؤشر تجزئة/تفادي موافقات ---
    near = [r for r in positive if threshold * 0.95 <= float(r["amount"]) < threshold]
    if near:
        findings.append(_finding(
            "LED_BELOW_THRESHOLD", "medium",
            "مبالغ قاربت سقف الصلاحية بقيمة أدنى منه",
            "Amounts just below the approval threshold",
            f"مبالغ تقع في نطاق {threshold*0.95:,.0f}–{threshold:,.0f} قد تدل على تفادي الحصول على الموافقات الأعلى (Structuring).",
            "Amounts clustered just under the threshold may indicate structuring to avoid higher-level approval.",
            "ISA 240 (fraud) / structuring",
            [{"date": r.get("date"), "beneficiary": r.get("beneficiary"), "amount": r.get("amount")} for r in near],
        ))

    # --- 8) تكرار أرقام الفواتير ---
    dup_inv = []
    if invoice_identifier:
        inv_buckets: dict[str, list[dict]] = {}
        for r in records:
            inv = clean_text(r.get("invoice_no")) or clean_text(r.get("auto_no"))
            if inv and re.search(r"\d", inv):
                inv_buckets.setdefault(inv, []).append(r)
        dup_inv = [
            {"invoice_no": k, "occurrences": len(v),
             "beneficiaries": sorted({clean_text(x.get("beneficiary")) for x in v})[:6],
             "amounts": [x.get("amount") for x in v][:6]}
            for k, v in inv_buckets.items() if len(v) > 1
        ]
    if dup_inv:
        findings.append(_finding(
            "LED_DUPLICATE_INVOICES", "medium",
            "تكرار أرقام فواتير",
            "Duplicate invoice numbers",
            "ظهور رقم فاتورة واحد أكثر من مرة قد يعني صرفاً مزدوجاً على الفاتورة نفسها.",
            "An invoice number used more than once may indicate double payment.",
            "ISA 240 (fraud)",
            dup_inv,
        ))

    # --- 9) القيم الشاذة إحصائياً (Outliers) ---
    if total >= 20:
        amounts = sorted(float(r["amount"]) for r in positive)

        def q(p):
            idx = min(len(amounts) - 1, max(0, int(round(p * (len(amounts) - 1)))))
            return amounts[idx]

        q1, q3 = q(0.25), q(0.75)
        hi = q3 + 3 * (q3 - q1)
        outliers = [r for r in positive if float(r["amount"]) > hi]
        if outliers:
            findings.append(_finding(
                "LED_STAT_OUTLIERS", "medium",
                "قيم مالية شاذة إحصائياً (Outliers)",
                "Statistically unusual amounts",
                "مبالغ تتجاوز الحد الإحصائي العلوي (Q3 + 3×IQR) وتستدعي فحصاً مستندياً معمّقاً.",
                "Amounts beyond the statistical upper fence (Q3 + 3×IQR) warrant detailed vouching.",
                "ISA 530 (Audit sampling) / analytical procedures",
                [{"date": r.get("date"), "beneficiary": r.get("beneficiary"), "amount": r.get("amount"), "check_no": r.get("check_no")} for r in outliers[:40]],
            ))

    # --- 10) تركّز المدفوعات لدى مستفيد واحد (Concentration) ---
    ben_totals: dict[str, float] = {}
    for r in positive:
        ben = clean_text(r.get("beneficiary"))
        if ben:
            ben_totals[ben] = ben_totals.get(ben, 0.0) + float(r["amount"])
    if total_amount > 0 and ben_totals:
        top = sorted(ben_totals.items(), key=lambda kv: kv[1], reverse=True)[:10]
        share = top[0][1] / total_amount * 100 if top else 0
        if share >= 40:
            findings.append(_finding(
                "LED_CONCENTRATION", "medium",
                "تركّز مرتفع للمدفوعات لدى مستفيد واحد",
                "High payment concentration in a single beneficiary",
                f"استحوذ المستفيد الأعلى على {share:.1f}% من إجمالي المدفوعات، ما يستدعي فحصاً لعلاقة المصالح وجودة الخدمة.",
                f"The top beneficiary received {share:.1f}% of total payments, warranting conflict-of-interest review.",
                "ISSAI 1240 / procurement governance",
                [{"beneficiary": b, "total": round(v, 2), "share_pct": round(v / total_amount * 100, 2)} for b, v in top],
            ))

    # --- 11) حركات نهاية السنة المالية ---
    year_end = []
    for r in records:
        d = to_date(r.get("date"))
        if d and d.month == 12 and d.day >= 25:
            year_end.append({"date": d.isoformat(), "beneficiary": r.get("beneficiary"), "amount": r.get("amount")})
    if len(year_end) >= 3:
        findings.append(_finding(
            "LED_YEAR_END", "low",
            "تركّز حركات في نهاية السنة المالية",
            "Clustering of transactions at fiscal year-end",
            "تركّز الصرف في الأيام الأخيرة من السنة قد يشير إلى تسريع تنفيذ مخصصات قبل الإغلاق.",
            "Year-end spending spikes may indicate rush to utilize appropriations before closing.",
            "ISSAI 1240 (cut-off) / ISA 560",
            year_end[:40],
        ))

    kpis = {
        "total_records": total,
        "total_amount": round(total_amount, 2),
        "average_amount": avg_amount,
        "max_amount": round(max((float(r["amount"]) for r in positive), default=0.0), 2),
        "duplicate_groups": len(dups),
        "split_groups": len(splits),
        "weekend_count": len(weekend),
        "negative_count": len(negatives),
        "distinct_beneficiaries": len({_norm_beneficiary(r.get("beneficiary") or "") for r in positive if _norm_beneficiary(r.get("beneficiary") or "")}),
    }
    details = {
        "duplicate_payments": dups,
        "split_purchases": splits,
        "weekend_transactions": weekend,
    }
    return findings, kpis, details


# ---------------------------------------------------------------------------
# المؤشرات والرسوم
# ---------------------------------------------------------------------------
def _position_kpis(items: list[dict]) -> dict:
    est = sum((it.get("estimated_budget") or 0) for it in items)
    paid = sum((it.get("total_paid") or 0) for it in items)
    ratios = [it.get("spend_ratio") for it in items if isinstance(it.get("spend_ratio"), (int, float))]
    util = round(paid / est * 100, 1) if est else 0.0
    return {
        "total_items": len(items),
        "total_budget": round(est, 2),
        "total_paid": round(paid, 2),
        "utilization_pct": util,
        "items_over_90": len([r for r in ratios if r >= 90]),
        "items_over_100": len([r for r in ratios if r >= 100]),
        "items_zero_movement": len([it for it in items if (it.get("estimated_budget") or 0) > 0 and not (it.get("total_paid") or 0) and not (it.get("financial_orders") or 0)]),
    }


def _position_charts(items: list[dict]) -> dict:
    ranked = [it for it in items if isinstance(it.get("spend_ratio"), (int, float))]
    ranked = sorted(ranked, key=lambda it: it["spend_ratio"], reverse=True)[:12]
    top_ratio = {
        "labels": [(it.get("item_name") or it.get("account_code") or "")[:40] for it in ranked],
        "values": [round(float(it["spend_ratio"]), 1) for it in ranked],
    }
    by_year: dict[int, float] = {}
    for it in items:
        y = it.get("year")
        if y:
            by_year[y] = by_year.get(y, 0) + (it.get("estimated_budget") or 0)
    years_sorted = sorted(by_year.keys())
    by_year_chart = {
        "labels": [str(y) for y in years_sorted],
        "values": [round(by_year[y], 2) for y in years_sorted],
    }
    return {"top_ratio": top_ratio, "by_year": by_year_chart}


def _ledger_charts(records: list[dict]) -> dict:
    monthly: dict[str, float] = {}
    for r in records:
        d = to_date(r.get("date"))
        if d and isinstance(r.get("amount"), (int, float)):
            key = f"{d.year}-{d.month:02d}"
            monthly[key] = monthly.get(key, 0.0) + float(r["amount"])
    months = sorted(monthly.keys())
    ben: dict[str, float] = {}
    for r in records:
        b = clean_text(r.get("beneficiary"))
        if b and isinstance(r.get("amount"), (int, float)):
            ben[b] = ben.get(b, 0.0) + float(r["amount"])
    top_ben = sorted(ben.items(), key=lambda kv: kv[1], reverse=True)[:8]
    return {
        "monthly": {"labels": months, "values": [round(monthly[m], 2) for m in months]},
        "top_beneficiaries": {
            "labels": [b[:35] for b, _ in top_ben],
            "values": [round(v, 2) for _, v in top_ben],
        },
    }


# ---------------------------------------------------------------------------
# الواجهة الرئيسية
# ---------------------------------------------------------------------------
def analyze_workbook(filename: str, contents: bytes) -> dict:
    """يحلّل ملفاً كاملاً ويعيد تقرير تدقيق مبني على المخاطر."""
    sheets = read_workbook(filename, contents)

    position_items: list[dict] = []
    ledger_records: list[dict] = []
    sheet_info: list[dict] = []
    ledger_fields: set[str] = set()

    for sheet_name, raw in sheets.items():
        if raw is None or raw.empty:
            sheet_info.append({"name": sheet_name, "type": "empty", "rows": 0})
            continue
        # إزالة الصفوف والأعمدة الفارغة تماماً
        raw = raw.dropna(axis=0, how="all").dropna(axis=1, how="all")
        if raw.empty:
            sheet_info.append({"name": sheet_name, "type": "empty", "rows": 0})
            continue
        raw = raw.reset_index(drop=True)
        # مزامنة عناوين الأعمدة مع مواقعها (بعد حذف الأعمدة الفارغة)
        raw.columns = list(range(raw.shape[1]))

        pos_map, _, pos_start = detect_columns(raw, POSITION_ALIASES)
        led_map, _, led_start = detect_columns(raw, LEDGER_ALIASES)
        kind = _sheet_kind(pos_map, led_map)

        if kind == "position":
            items = extract_position_items(raw, pos_map, pos_start, sheet_name)
            items = [it for it in items if (it.get("estimated_budget") or it.get("total_paid") or it.get("financial_orders"))]
            position_items.extend(items)
            sheet_info.append({"name": sheet_name, "type": "position", "rows": len(items)})
        elif kind == "ledger":
            recs = extract_ledger_records(raw, led_map, led_start, sheet_name)
            ledger_records.extend(recs)
            ledger_fields |= set(led_map.values())
            sheet_info.append({"name": sheet_name, "type": "ledger", "rows": len(recs)})
        else:
            sheet_info.append({"name": sheet_name, "type": "unknown", "rows": 0})

    findings: list[dict] = []
    details: dict[str, Any] = {}
    kpis: dict[str, Any] = {}
    charts: dict[str, Any] = {}
    file_type = "unknown"

    pos_kpis = _position_kpis(position_items) if position_items else {}
    led_findings: list[dict] = []
    led_kpis: dict[str, Any] = {}
    led_details: dict[str, Any] = {}

    if position_items:
        findings.extend(_audit_position(position_items, pos_kpis))
        kpis.update(pos_kpis)
        charts["position"] = _position_charts(position_items)
    if ledger_records:
        invoice_identifier = ("invoice_no" in ledger_fields) or ("invoice_type" in ledger_fields)
        led_findings, led_kpis, led_details = _audit_ledger(ledger_records, invoice_identifier=invoice_identifier)
        findings.extend(led_findings)
        kpis.update(led_kpis)
        charts["ledger"] = _ledger_charts(ledger_records)
        details.update(led_details)

    if position_items and ledger_records:
        file_type = "mixed"
    elif position_items:
        file_type = "financial_position"
    elif ledger_records:
        file_type = "ledger"

    # فرز الملاحظات حسب الخطورة ثم الحجم
    order = {"high": 0, "medium": 1, "low": 2}
    findings.sort(key=lambda f: (order.get(f["severity"], 3), -f.get("count", 0)))

    risk_score = _risk_score(findings, len(position_items) + len(ledger_records))
    risk_level = "high" if risk_score >= 60 else ("medium" if risk_score >= 30 else "low")

    total_scanned = len(position_items) + len(ledger_records)
    summary: dict[str, Any] = {"file_type": file_type}
    for f in findings:
        summary[f["id"]] = f["count"]
    summary["high_count"] = len([f for f in findings if f["severity"] == "high"])
    summary["medium_count"] = len([f for f in findings if f["severity"] == "medium"])
    summary["low_count"] = len([f for f in findings if f["severity"] == "low"])

    return {
        "filename": filename,
        "file_type": file_type,
        "sheet_info": sheet_info,
        "total_records_scanned": total_scanned,
        "risk_score": risk_score,
        "risk_level": risk_level,
        "summary": summary,
        "kpis": kpis,
        "charts": charts,
        "findings": findings,
        "position_preview": position_items[:200],
        "ledger_preview": ledger_records[:200],
        "details": details,
    }


def _risk_score(findings: list[dict], total_records: int) -> int:
    if not findings:
        return 0
    raw = 0.0
    for f in findings:
        cnt = max(1, int(f.get("count", 1)))
        raw += SEV_WEIGHT.get(f["severity"], 1.0) * (1 + math.log10(cnt) / 2.0)
    # معايرة: كلما زاد عدد السجلات قل تأثير النسبة، لكن نبقي الحد الأدنى
    score = min(100, round(raw))
    if findings and score < 8:
        score = 8
    return int(score)


# ---------------------------------------------------------------------------
# ملخّص نصي للذكاء الاصطناعي
# ---------------------------------------------------------------------------
def build_digest(report: dict, lang: str = "ar", max_items: int = 8) -> str:
    import json as _json
    lines: list[str] = []
    if lang == "en":
        lines.append(f"File: {report.get('filename')}")
        lines.append(f"File type: {report.get('file_type')}")
        lines.append(f"Records analysed: {report.get('total_records_scanned')}")
        lines.append(f"Overall risk score: {report.get('risk_score')}% ({report.get('risk_level')})")
        lines.append(f"KPIs: {_json.dumps(report.get('kpis', {}), ensure_ascii=False)}")
    else:
        lines.append(f"الملف: {report.get('filename')}")
        lines.append(f"نوع الملف: {report.get('file_type')}")
        lines.append(f"عدد السجلات المحللة: {report.get('total_records_scanned')}")
        lines.append(f"درجة الخطورة العامة: {report.get('risk_score')}% ({report.get('risk_level')})")
        lines.append(f"المؤشرات: {_json.dumps(report.get('kpis', {}), ensure_ascii=False)}")
    lines.append("")
    lines.append("Findings / الملاحظات:" if lang == "en" else "الملاحظات المكتشفة:")
    for f in report.get("findings", []):
        title = f["title_en"] if lang == "en" else f["title_ar"]
        sev = f["severity_en"] if lang == "en" else f["severity_ar"]
        lines.append(f"- [{sev}] {title} — count={f['count']} (standard: {f['standard']})")
        for it in (f.get("items") or [])[:max_items]:
            lines.append("    • " + _json.dumps(it, ensure_ascii=False, default=str))
    return "\n".join(lines)


def build_chat_context(report: dict, lang: str = "ar") -> str:
    """سياق كامل نسبياً للدردشة التفاعلية."""
    return build_digest(report, lang=lang, max_items=15)
