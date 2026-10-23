"""
نظام تدقيق البيانات الحكومية — Government Data Auditing System
================================================================
Backend  : FastAPI + Pandas + Ollama (Local AI)
Features : CSV/Excel auditing, configurable Ollama endpoint,
           bilingual (Arabic/English) AI analysis, dashboard data.
"""

from fastapi import FastAPI, UploadFile, File, Body
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import FileResponse, JSONResponse, StreamingResponse
from fastapi.staticfiles import StaticFiles
import pandas as pd
import io
import os
import re
import sys
import json
import html
import atexit
import hashlib
import threading
from datetime import datetime
import requests
import socket

import audit_engine as engine

# ---------------------------------------------------------------------------
# Force UTF-8 console output.
# On Windows the default console code page (cp1252/cp720) cannot encode Arabic
# and raises UnicodeEncodeError, which would crash the app on startup.
# ---------------------------------------------------------------------------
for _stream in (sys.stdout, sys.stderr):
    try:
        _stream.reconfigure(encoding="utf-8")  # Python 3.7+
    except Exception:
        pass


# ---------------------------------------------------------------------------
# Paths & configuration
# ---------------------------------------------------------------------------
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
INDEX_PATH = os.path.join(BASE_DIR, "index.html")
CONFIG_PATH = os.path.join(BASE_DIR, "config.json")
LOGO_PATH = os.path.join(BASE_DIR, "MAIN DATA", "LOGOAB.jfif")

DEFAULT_CONFIG = {
    "ollama_url": "http://127.0.0.1:11434",  # can point to any machine/device
    "model": "qwen3:4b",
    "language": "ar",              # ar | en
    "max_ai_records": 40,          # max anomalies per rule sent to the model
    "temperature": 0.2,
    "num_predict": 4096,           # max tokens generated per answer
    "request_timeout": 300,        # seconds — local models can be slow
    "think": False,                # set true to let reasoning models (qwen3, ...) think
}


def load_config() -> dict:
    """Read config.json (falling back to defaults for missing keys)."""
    cfg = dict(DEFAULT_CONFIG)
    try:
        with open(CONFIG_PATH, "r", encoding="utf-8") as fh:
            cfg.update(json.load(fh))
    except Exception:
        pass
    return cfg


def save_config(new: dict) -> dict:
    """Persist only known keys so the file stays clean."""
    cfg = load_config()
    for key in DEFAULT_CONFIG:
        if isinstance(new, dict) and new.get(key) is not None:
            cfg[key] = new[key]
    with open(CONFIG_PATH, "w", encoding="utf-8") as fh:
        json.dump(cfg, fh, ensure_ascii=False, indent=2)
    return cfg


# ---------------------------------------------------------------------------
# Application state (app_state.json)
# ---------------------------------------------------------------------------
# A single JSON file written to the program folder EVERY time the program
# updates (startup, shutdown, settings change, file uploaded, audit run, AI
# analysis, UI navigation, ...). It always carries a fresh timestamp so the
# current/previous state of the application can be inspected at any time.
STATE_PATH = os.path.join(BASE_DIR, "app_state.json")
APP_NAME_AR = "المدقق الذكي"
APP_NAME_EN = "Smart Auditor"
APP_VERSION = "2.0"
_state_lock = threading.Lock()
_STATE_HISTORY_LIMIT = 300


def _now_str() -> str:
    return datetime.now().strftime("%Y-%m-%d %H:%M:%S")


def read_app_state() -> dict:
    """Return the current app_state.json (empty dict if none yet)."""
    try:
        with open(STATE_PATH, "r", encoding="utf-8") as fh:
            return json.load(fh)
    except Exception:
        return {}


def _write_app_state(state: dict) -> None:
    """Atomic write (temp file + replace) so a crash can never corrupt it."""
    tmp = STATE_PATH + ".tmp"
    with open(tmp, "w", encoding="utf-8") as fh:
        json.dump(state, fh, ensure_ascii=False, indent=2)
    os.replace(tmp, STATE_PATH)


def _clean_details(details: dict) -> dict:
    """Keep detail values JSON-safe and short so the state file stays clean."""
    out = {}
    for key, value in (details or {}).items():
        if isinstance(value, str) and len(value) > 400:
            value = value[:400] + "…"
        else:
            try:
                json.dumps(value)
            except Exception:
                value = str(value)
        out[key] = value
    return out


def update_app_state(event: str, _inc: dict | None = None, **details) -> dict:
    """Create/refresh app_state.json with a timestamp on every program update."""
    with _state_lock:
        state = read_app_state()
        now = _now_str()
        if not state:
            state = {
                "app_name": APP_NAME_AR,
                "app_name_en": APP_NAME_EN,
                "version": APP_VERSION,
                "state_file": os.path.basename(STATE_PATH),
                "created_at": now,
                "update_count": 0,
                "stats": {"audits": 0, "analyses": 0},
                "history": [],
            }
        state["app_name"] = APP_NAME_AR
        state["app_name_en"] = APP_NAME_EN
        state["version"] = APP_VERSION
        state["updated_at"] = now
        state["updated_at_iso"] = datetime.now().isoformat(timespec="seconds")
        state["update_count"] = int(state.get("update_count", 0)) + 1
        if _inc:
            stats = state.setdefault("stats", {})
            for key, value in _inc.items():
                stats[key] = int(stats.get(key, 0)) + int(value)
        last = {"time": now, "type": event, "details": _clean_details(details)}
        state["last_event"] = last
        history = state.setdefault("history", [])
        history.append(last)
        if len(history) > _STATE_HISTORY_LIMIT:
            del history[:-_STATE_HISTORY_LIMIT]
        try:
            state["config"] = load_config()
        except Exception:
            pass
        _write_app_state(state)
        return state


# First creation of the state file as soon as the program starts.
if not os.path.exists(STATE_PATH):
    try:
        update_app_state("install", note="app_state.json created")
    except Exception:
        pass


# ---------------------------------------------------------------------------
# App
# ---------------------------------------------------------------------------
from contextlib import asynccontextmanager


@asynccontextmanager
async def _lifespan(_app):
    # Log the active event loop at startup — useful to confirm the Windows
    # Selector loop is in use (the Proactor loop's accept loop can die, WinError 64).
    try:
        import asyncio
        print(f"[diag] running event loop type={type(asyncio.get_running_loop()).__name__}", flush=True)
    except Exception as exc:  # noqa: BLE001
        print(f"[diag] {exc}", flush=True)
    yield


app = FastAPI(title="نظام تدقيق البيانات الحكومية", version="2.0", lifespan=_lifespan)

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=False,  # wildcard origin + credentials is invalid per CORS
    allow_methods=["*"],
    allow_headers=["*"],
)


# ---------------------------------------------------------------------------
# In-memory store of recent audit reports (used by AI analysis & chat).
# Keyed by a short report id; only the latest reports are kept.
# ---------------------------------------------------------------------------
REPORTS: dict[str, dict] = {}
REPORTS_LOCK = threading.Lock()
REPORTS_LIMIT = 20


def _store_report(report: dict) -> str:
    filename = report.get("filename", "file")
    stamp = datetime.now().strftime("%Y%m%d%H%M%S%f")
    rid = hashlib.md5(f"{filename}{stamp}".encode("utf-8")).hexdigest()[:16]
    with REPORTS_LOCK:
        REPORTS[rid] = report
        if len(REPORTS) > REPORTS_LIMIT:
            for key in list(REPORTS.keys())[:-REPORTS_LIMIT]:
                REPORTS.pop(key, None)
    return rid


# ---------------------------------------------------------------------------
# Static frontend
# ---------------------------------------------------------------------------
@app.get("/", include_in_schema=False)
def index():
    if os.path.exists(INDEX_PATH):
        return FileResponse(INDEX_PATH)
    return JSONResponse({"error": "index.html not found"}, status_code=404)


@app.get("/logo", include_in_schema=False)
def logo():
    """Company logo (served so the UI shows it without exposing the folder)."""
    if os.path.exists(LOGO_PATH):
        return FileResponse(LOGO_PATH, media_type="image/jpeg")
    return JSONResponse({"error": "logo_not_found"}, status_code=404)


# Local static assets (Bootstrap / FontAwesome / Cairo / Chart.js) so the UI
# works 100% offline — no CDN requests, no internet required.
STATIC_DIR = os.path.join(BASE_DIR, "static")
if os.path.isdir(STATIC_DIR):
    app.mount("/static", StaticFiles(directory=STATIC_DIR), name="static")


def _lan_addresses() -> list[str]:
    ips: set[str] = set()
    try:
        hostname = socket.gethostname()
        for info in socket.getaddrinfo(hostname, socket.AF_INET):
            ip = info[4][0]
            if ip and not ip.startswith("127.") and "." in ip:
                ips.add(ip)
    except Exception:
        pass
    try:
        s = socket.socket(socket.AF_INET, socket.SOCK_DGRAM)
        s.settimeout(0.5)
        s.connect(("8.8.8.8", 80))
        ip = s.getsockname()[0]
        s.close()
        if ip and not ip.startswith("127."):
            ips.add(ip)
    except Exception:
        pass
    return sorted(ips)


@app.get("/health")
def health():
    return {"status": "ok", "service": "government-audit", "version": "2.0",
            "lan_addresses": _lan_addresses()}


# ---------------------------------------------------------------------------
# Application state (app_state.json) — readable from the UI / scripts
# ---------------------------------------------------------------------------
@app.get("/state")
def get_app_state():
    """Return the current application state (timestamped on every update)."""
    return read_app_state()


@app.post("/state/event")
def post_app_state_event(payload: dict = Body(...)):
    """Record a UI/application event (language change, navigation, ...)."""
    event = str(payload.get("type") or "event")
    details = payload.get("details") or {}
    if not isinstance(details, dict):
        details = {"value": details}
    state = update_app_state("ui_" + event, **details)
    return {"ok": True, "updated_at": state.get("updated_at"), "update_count": state.get("update_count")}


# ---------------------------------------------------------------------------
# Settings (Ollama endpoint is fully configurable -> portable between devices)
# ---------------------------------------------------------------------------
@app.get("/settings")
def get_settings():
    return load_config()


@app.post("/settings")
def update_settings(payload: dict = Body(...)):
    cfg = save_config(payload)
    update_app_state(
        "settings_update",
        ollama_url=cfg.get("ollama_url"),
        model=cfg.get("model"),
        language=cfg.get("language"),
    )
    return cfg


# ---------------------------------------------------------------------------
# Ollama helpers
# ---------------------------------------------------------------------------
@app.get("/ollama/models")
def ollama_models(url: str | None = None):
    """List models available on the configured (or overridden) Ollama server."""
    base = (url or load_config().get("ollama_url") or "").rstrip("/")
    if not base:
        return {"ok": False, "models": [], "error": "no_url"}
    try:
        resp = requests.get(f"{base}/api/tags", timeout=8)
        resp.raise_for_status()
        models = [m.get("name") for m in resp.json().get("models", []) if m.get("name")]
        update_app_state("models_refresh", ok=True, url=base, count=len(models))
        return {"ok": True, "url": base, "models": sorted(models), "count": len(models)}
    except Exception as exc:  # noqa: BLE001 - surface any connection problem to UI
        update_app_state("models_refresh", ok=False, url=base, error=str(exc))
        return {"ok": False, "url": base, "models": [], "error": str(exc)}


def _strip_think(text: str) -> str:
    """Remove qwen3-style  thinking...<｜end▁of▁thinking｜> reasoning blocks from text."""
    return re.sub(r"(?is) thinking.*?<｜end▁of▁thinking｜>", "", text or "").strip()


def _ollama_body(messages: list, model: str, cfg: dict, stream: bool) -> dict:
    body = {
        "model": model,
        "messages": messages,
        "stream": stream,
        "options": {
            "temperature": float(cfg.get("temperature", 0.2)),
            "num_predict": int(cfg.get("num_predict", 1500)),
        },
    }
    # Reasoning models (qwen3, deepseek-r1, ...) can spend most of their budget on
    # hidden "thinking" tokens, which makes interactive chat feel frozen. Ollama
    # 0.9+ accepts a top-level `think` flag; other models simply ignore it.
    body["think"] = bool(cfg.get("think", False))
    return body


def _iter_ollama(resp) -> "generator":
    """Translate Ollama's NDJSON stream into typed events (thinking/content/done).

    Never raises: any mid-stream failure is reported as an 'error' event so a
    broken model or a dropped connection can never take the worker down.
    """
    try:
        try:
            for line in resp.iter_lines(decode_unicode=True):
                if not line:
                    continue
                try:
                    obj = json.loads(line)
                except Exception:
                    continue
                msg = obj.get("message") or {}
                if msg.get("thinking"):
                    yield json.dumps({"type": "thinking", "text": msg["thinking"]}, ensure_ascii=False) + "\n"
                if msg.get("content"):
                    yield json.dumps({"type": "content", "text": msg["content"]}, ensure_ascii=False) + "\n"
                if obj.get("error"):
                    yield json.dumps({"type": "error", "message": obj["error"]}, ensure_ascii=False) + "\n"
                if obj.get("done"):
                    break
        except Exception as exc:  # noqa: BLE001 - stream aborted / connection lost
            yield json.dumps({"type": "error", "message": str(exc)}, ensure_ascii=False) + "\n"
    finally:
        try:
            resp.close()
        except Exception:
            pass
    yield json.dumps({"type": "done"}) + "\n"


def _classify_error(exc: Exception) -> str:
    if isinstance(exc, requests.exceptions.ConnectTimeout):
        return "timeout"
    if isinstance(exc, requests.exceptions.ConnectionError):
        return "connection"
    if isinstance(exc, requests.exceptions.HTTPError):
        return "http"
    return "unknown"


def _audit_system_prompt(lang: str) -> str:
    if lang == "en":
        return (
            "You are a senior government financial auditor (Supreme Audit Institution) "
            "specialised in risk-based auditing, fraud detection and regularity/compliance "
            "auditing, following international standards (INTOSAI ISSAI, IFAC ISA 240/315/1260). "
            "You MUST rely exclusively on the numbers and findings provided — never invent figures. "
            "Reply ONLY in English, in clear Markdown, with these sections: "
            "## Executive summary, ## Scope & methodology, ## Key findings (by risk), "
            "## Risk assessment (Low/Medium/High), ## Compliance & regularity observations, "
            "## Recommendations & corrective actions. Be precise, professional and actionable."
        )
    return (
        "أنت مدقق مالي حكومي أول في ديوان محاسبة/جهاز رقابي، متخصص في التدقيق المبني على المخاطر "
        "وكشف الاحتيال ومخالفات الانضباط المالي وفق المعايير الدولية (INTOSAI ISSAI، وIFAC ISA 240/315/1260). "
        "يجب أن تعتمد حصرياً على الأرقام والملاحظات المرفقة ولا تختلق أي رقم. "
        "اكتب ردك باللغة العربية فقط، بتنسيق Markdown واضح، وفق الأقسام التالية: "
        "## الملخص التنفيذي، ## نطاق التدقيق والمنهجية، ## أبرز الملاحظات (مرتبة حسب الخطورة)، "
        "## تقييم المخاطر (منخفضة/متوسطة/عالية)، ## ملاحظات الالتزام والانضباط المالي، "
        "## التوصيات وإجراءات التصحيح. كن دقيقاً ومهنياً وقابلاً للتنفيذ."
    )


def _prompt_messages(lang: str, digest: str) -> list:
    user = (
        f"Audit findings to analyse:\n\n{digest}"
        if lang == "en"
        else f"نتائج التدقيق المطلوب تحليلها:\n\n{digest}"
    )
    return [
        {"role": "system", "content": _audit_system_prompt(lang)},
        {"role": "user", "content": user},
    ]


def _chat_system_prompt(lang: str, context: str) -> str:
    if lang == "en":
        return (
            "You are a senior government financial auditor answering questions about a specific "
            "audit file. Use ONLY the audit context below (real numbers from the file). "
            "If a figure is not present, say so explicitly — never fabricate. "
            "Answer concisely in English with Markdown, cite the relevant item/account/beneficiary.\n\n"
            f"--- AUDIT CONTEXT ---\n{context}"
        )
    return (
        "أنت مدقق مالي حكومي أول تجيب على أسئلة المدقق حول ملف تدقيق محدد. "
        "اعتمد حصرياً على سياق التدقيق أدناه (أرقام حقيقية من الملف). "
        "إذا لم يتوفر رقم أو بند فاذكر ذلك صراحةً ولا تختلق أي معلومة. "
        "أجب بإيجاز ودقة باللغة العربية بتنسيق Markdown مع الإشارة إلى البند/الحساب/المستفيد المعني.\n\n"
        f"--- سياق التدقيق ---\n{context}"
    )


def _resolve_report(payload: dict) -> dict | None:
    report = payload.get("report")
    if isinstance(report, dict) and report.get("findings") is not None:
        return report
    rid = payload.get("report_id")
    if rid:
        with REPORTS_LOCK:
            return REPORTS.get(rid)
    return None


# ---------------------------------------------------------------------------
# AI analysis (Local LLM via Ollama)
# ---------------------------------------------------------------------------
@app.post("/analyze/")
def analyze(payload: dict = Body(...)):
    """Stream a professional AI audit report generated from the audit findings."""
    cfg = load_config()
    lang = "en" if str(payload.get("language", cfg.get("language", "ar"))).lower().startswith("en") else "ar"
    model = payload.get("model") or cfg.get("model")
    url = payload.get("ollama_url") or cfg.get("ollama_url")

    if not model:
        return JSONResponse({"ok": False, "error": "no_model"})
    if not url:
        return JSONResponse({"ok": False, "error": "no_url"})

    report = _resolve_report(payload)
    if report is None:
        return JSONResponse({"ok": False, "error": "no_report"})

    limit = int(cfg.get("max_ai_records", 40))
    digest = engine.build_digest(report, lang=lang, max_items=max(3, limit // 6))
    messages = _prompt_messages(lang, digest)
    base = url.rstrip("/")
    body = _ollama_body(messages, model, cfg, stream=True)

    update_app_state("analyze_start", model=model, language=lang, url=base)

    # Pre-flight the connection so real errors surface as JSON, not mid-stream.
    try:
        resp = requests.post(
            f"{base}/api/chat",
            json=body,
            stream=True,
            timeout=(10, int(cfg.get("request_timeout", 300))),
        )
        resp.raise_for_status()
    except Exception as exc:  # noqa: BLE001
        update_app_state("analyze_error", model=model, error=_classify_error(exc))
        return JSONResponse({"ok": False, "error": _classify_error(exc), "detail": str(exc)})

    def event_stream():
        yield json.dumps({"type": "meta", "model": model, "language": lang},
                         ensure_ascii=False) + "\n"
        yield from _iter_ollama(resp)
        update_app_state("analyze_done", _inc={"analyses": 1}, model=model, language=lang)

    return StreamingResponse(event_stream(), media_type="application/x-ndjson")


@app.post("/chat/")
def chat(payload: dict = Body(...)):
    """Interactive audit chatbot grounded strictly in the uploaded file's numbers."""
    cfg = load_config()
    lang = "en" if str(payload.get("language", cfg.get("language", "ar"))).lower().startswith("en") else "ar"
    model = payload.get("model") or cfg.get("model")
    url = payload.get("ollama_url") or cfg.get("ollama_url")
    question = str(payload.get("question") or "").strip()

    if not model:
        return JSONResponse({"ok": False, "error": "no_model"})
    if not url:
        return JSONResponse({"ok": False, "error": "no_url"})
    if not question:
        return JSONResponse({"ok": False, "error": "no_question"})

    report = _resolve_report(payload)
    if report is None:
        return JSONResponse({"ok": False, "error": "no_report"})

    context = engine.build_chat_context(report, lang=lang)
    messages = [{"role": "system", "content": _chat_system_prompt(lang, context)}]
    for turn in (payload.get("history") or [])[-8:]:
        role = turn.get("role")
        content = turn.get("content")
        if role in ("user", "assistant") and content:
            messages.append({"role": role, "content": str(content)})
    messages.append({"role": "user", "content": question})

    base = url.rstrip("/")
    body = _ollama_body(messages, model, cfg, stream=True)
    update_app_state("chat_start", model=model, language=lang)

    try:
        resp = requests.post(
            f"{base}/api/chat",
            json=body,
            stream=True,
            timeout=(10, int(cfg.get("request_timeout", 300))),
        )
        resp.raise_for_status()
    except Exception as exc:  # noqa: BLE001
        update_app_state("chat_error", model=model, error=_classify_error(exc))
        return JSONResponse({"ok": False, "error": _classify_error(exc), "detail": str(exc)})

    def event_stream():
        yield json.dumps({"type": "meta", "model": model, "language": lang},
                         ensure_ascii=False) + "\n"
        yield from _iter_ollama(resp)
        update_app_state("chat_done", _inc={"chats": 1}, model=model)

    return StreamingResponse(event_stream(), media_type="application/x-ndjson")


# ---------------------------------------------------------------------------
# Auditing engine — flexible multi-sheet reader + risk-based rules
# ---------------------------------------------------------------------------
@app.post("/audit/")
async def audit_data(file: UploadFile = File(...)):
    contents = await file.read()
    try:
        report = engine.analyze_workbook(file.filename, contents)
    except ValueError as exc:
        reason = str(exc)
        if reason == "unsupported_format":
            update_app_state("audit_error", filename=file.filename, reason="unsupported_format")
            return {"error": "unsupported_format"}
        update_app_state("audit_error", filename=file.filename, reason="read_failed", detail=reason)
        return {"error": "read_failed", "detail": reason}
    except Exception as exc:  # noqa: BLE001 - corrupted Excel / missing engine
        update_app_state("audit_error", filename=file.filename, reason="read_failed", detail=str(exc))
        return {"error": "read_failed", "detail": str(exc)}

    if report.get("total_records_scanned", 0) == 0:
        update_app_state("audit_error", filename=file.filename, reason="no_recognised_data")
        return {"error": "no_recognised_data"}

    report_id = _store_report(report)
    report["report_id"] = report_id
    update_app_state(
        "audit_run",
        _inc={"audits": 1},
        filename=file.filename,
        records=report.get("total_records_scanned"),
        risk=report.get("risk_score"),
        findings={k: v for k, v in report.get("summary", {}).items() if isinstance(v, int)},
    )
    return report


# ---------------------------------------------------------------------------
# Report export (self-contained printable HTML — also usable as PDF via print)
# ---------------------------------------------------------------------------
def _report_html(report: dict, lang: str) -> str:
    rtl = lang != "en"
    esc = html.escape

    T = {
        "title": "تقرير التدقيق المالي المبني على المخاطر" if rtl else "Risk-Based Financial Audit Report",
        "file": "الملف" if rtl else "File",
        "type": "نوع الملف" if rtl else "File type",
        "records": "السجلات المحللة" if rtl else "Records analysed",
        "risk": "درجة الخطورة" if rtl else "Risk score",
        "level": "المستوى" if rtl else "Level",
        "kpis": "المؤشرات الرئيسية" if rtl else "Key indicators",
        "findings": "الملاحظات والمخالفات" if rtl else "Findings & violations",
        "count": "العدد" if rtl else "Count",
        "standard": "المعيار المرجعي" if rtl else "Reference standard",
        "generated": "تاريخ الإصدار" if rtl else "Generated",
        "no_findings": "لا توجد ملاحظات." if rtl else "No findings.",
        "sev": {"high": "عالية" if rtl else "High", "medium": "متوسطة" if rtl else "Medium", "low": "منخفضة" if rtl else "Low"},
        "type_map": {
            "financial_position": "مواقف مالية" if rtl else "Financial position",
            "ledger": "سجل حركات" if rtl else "Transaction ledger",
            "mixed": "مواقف مالية + سجلات" if rtl else "Position + ledger",
        },
    }

    def fmt(v):
        if isinstance(v, float):
            return f"{v:,.2f}"
        if isinstance(v, int):
            return f"{v:,}"
        if v is None:
            return ""
        return esc(str(v))

    sev_color = {"high": "#dc2626", "medium": "#d97706", "low": "#2563eb"}
    parts: list[str] = []
    parts.append("<!DOCTYPE html><html lang='%s' dir='%s'><head><meta charset='utf-8'>" % (lang, "rtl" if rtl else "ltr"))
    parts.append("<title>%s</title><style>" % esc(T["title"]))
    parts.append(
        "body{font-family:'Segoe UI',Tahoma,Arial,sans-serif;margin:32px;color:#1e293b;background:#fff}"
        "h1{font-size:22px}h2{font-size:17px;border-bottom:2px solid #e2e8f0;padding-bottom:6px;margin-top:30px}"
        ".meta{display:flex;flex-wrap:wrap;gap:18px;background:#f8fafc;border:1px solid #e2e8f0;border-radius:10px;padding:14px}"
        ".meta div{font-size:13px}.meta b{color:#0f172a}"
        ".risk{padding:6px 14px;border-radius:8px;color:#fff;font-weight:700}"
        "table{width:100%;border-collapse:collapse;margin:8px 0 18px;font-size:12px}"
        "th,td{border:1px solid #e2e8f0;padding:6px 8px;text-align:start}"
        "th{background:#f1f5f9}"
        ".badge{display:inline-block;padding:2px 8px;border-radius:6px;color:#fff;font-size:11px}"
        ".kpi{display:inline-block;min-width:150px;background:#f8fafc;border:1px solid #e2e8f0;border-radius:10px;padding:10px 14px;margin:0 10px 10px 0}"
        ".kpi b{display:block;font-size:18px}"
    )
    parts.append("</style></head><body>")
    parts.append(f"<h1>{esc(T['title'])}</h1>")
    risk = report.get("risk_score", 0)
    level = report.get("risk_level", "low")
    parts.append("<div class='meta'>")
    parts.append(f"<div><b>{T['file']}:</b> {esc(str(report.get('filename','')))}</div>")
    parts.append(f"<div><b>{T['type']}:</b> {esc(T['type_map'].get(report.get('file_type',''), report.get('file_type','')))}</div>")
    parts.append(f"<div><b>{T['records']}:</b> {fmt(report.get('total_records_scanned', 0))}</div>")
    parts.append(f"<div><b>{T['risk']}:</b> {risk}% </div>")
    parts.append(f"<div><span class='risk' style='background:{sev_color.get(level,'#334155')}'>{T['sev'].get(level, level)}</span></div>")
    parts.append(f"<div><b>{T['generated']}:</b> {datetime.now().strftime('%Y-%m-%d %H:%M')}</div>")
    parts.append("</div>")

    kpis = report.get("kpis", {}) or {}
    if kpis:
        parts.append(f"<h2>{esc(T['kpis'])}</h2>")
        for k, v in kpis.items():
            parts.append(f"<span class='kpi'>{esc(str(k))}<b>{fmt(v)}</b></span>")

    findings = report.get("findings", []) or []
    parts.append(f"<h2>{esc(T['findings'])}</h2>")
    if not findings:
        parts.append(f"<p>{esc(T['no_findings'])}</p>")
    for f in findings:
        title = f.get("title_en") if not rtl else f.get("title_ar")
        desc = f.get("desc_en") if not rtl else f.get("desc_ar")
        sev = f.get("severity", "low")
        parts.append(
            f"<h3><span class='badge' style='background:{sev_color.get(sev,'#334155')}'>"
            f"{T['sev'].get(sev, sev)}</span> {esc(str(title))} "
            f"<span style='font-weight:400;color:#64748b'>— {esc(T['count'])}: {f.get('count',0)}</span></h3>"
        )
        if desc:
            parts.append(f"<p style='font-size:13px;color:#334155'>{esc(str(desc))}</p>")
        parts.append(f"<p style='font-size:12px;color:#64748b'><b>{T['standard']}:</b> {esc(str(f.get('standard','')))}</p>")
        items = f.get("items") or []
        if items:
            keys: list[str] = []
            for it in items:
                if isinstance(it, dict):
                    for k in it.keys():
                        if k not in keys:
                            keys.append(k)
            parts.append("<table><thead><tr>" + "".join(f"<th>{esc(str(k))}</th>" for k in keys) + "</tr></thead><tbody>")
            for it in items:
                parts.append("<tr>" + "".join(f"<td>{fmt(it.get(k) if isinstance(it, dict) else '')}</td>" for k in keys) + "</tr>")
            parts.append("</tbody></table>")
    parts.append("</body></html>")
    return "".join(parts)


@app.post("/export/")
def export_report(payload: dict = Body(...)):
    report = _resolve_report(payload)
    if report is None:
        return JSONResponse({"ok": False, "error": "no_report"}, status_code=404)
    lang = "en" if str(payload.get("language", "ar")).lower().startswith("en") else "ar"
    html_doc = _report_html(report, lang)
    filename = f"audit_report_{datetime.now().strftime('%Y%m%d_%H%M')}.html"
    return StreamingResponse(
        io.BytesIO(html_doc.encode("utf-8")),
        media_type="text/html; charset=utf-8",
        headers={
            "Content-Disposition": f'attachment; filename="{filename}"',
            "Access-Control-Expose-Headers": "Content-Disposition",
        },
    )


if __name__ == "__main__":
    import asyncio

    # uvicorn >=0.36 *forces* the Windows Proactor event loop (asyncio_loop_factory
    # returns ProactorEventLoop on win32 for a single worker), ignoring any policy
    # set here. The Proactor accept loop can then die permanently (OSError WinError
    # 64 "network name is no longer available") after a client abruptly resets a
    # connection — leaving the server process alive but no longer listening, which
    # the UI reports as "تعذّر الاتصال بالسيرفر". We therefore hand uvicorn an
    # explicit loop factory that builds the stable Selector loop (this app is
    # HTTP-only and spawns no subprocesses).
    def _selector_loop_factory() -> "asyncio.AbstractEventLoop":
        return asyncio.SelectorEventLoop()

    import uvicorn

    cfg = load_config()
    print("=" * 62, flush=True)
    print("  Government Data Auditing System  |  نظام تدقيق البيانات", flush=True)
    print("  Open in browser :  http://127.0.0.1:8000", flush=True)
    _lan = _lan_addresses()
    for ip in _lan:
        print(f"  LAN (local network) :  http://{ip}:8000", flush=True)
    print(f"  gemini endpoint :  {cfg.get('ollama_url')}", flush=True)
    print(f"  Default model   :  {cfg.get('model')}", flush=True)
    print("  Press Ctrl+C to stop.", flush=True)
    print("=" * 62, flush=True)

    update_app_state("server_started", host="0.0.0.0", port=8000, pid=os.getpid(), lan_urls=_lan_addresses())
    atexit.register(lambda: update_app_state("server_stopped", pid=os.getpid()))

    if os.name == "nt":
        asyncio.set_event_loop_policy(asyncio.WindowsSelectorEventLoopPolicy())
        uvicorn.run(app, host="0.0.0.0", port=8000, loop="__main__:_selector_loop_factory")
    else:
        uvicorn.run(app, host="0.0.0.0", port=8000)
