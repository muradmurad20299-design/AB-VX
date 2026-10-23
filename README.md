# المدقق الذكي — Smart Auditor

نظام تدقيق مالي حكومي مبني على المخاطر، يقرأ ملفات Excel/CSV بأسماء أعمدة وترتيب ولغة متغيّرة
(عربي/إنجليزي) ويكتشف المخالفات ومؤشرات الاحتيال وفق المعايير الدولية، مع تحليل ذكي محلي عبر Ollama.

**A risk‑based government financial‑audit system.** It reads Excel/CSV files with flexible
column names, order and language (Arabic/English), detects irregularities and fraud indicators
aligned with international standards, and adds a local AI analysis via Ollama.

---

## 1) المتطلبات — Requirements

- Python 3.10+ (تم الاختبار على 3.12)
- [Ollama](https://ollama.com) يعمل محلياً على `http://127.0.0.1:11434` (اختياري للتحليل الذكي)
- حِزم بايثون في `requirements.txt`:
  `fastapi`, `uvicorn`, `pandas`, `requests`, `openpyxl`, `xlrd`, `pyxlsb`, `python-multipart`

> `pyxlsb` مطلوب لقراءة ملفات `.xlsb` (مثل «سجل الفواتير 2025.xlsb»).

## 2) التثبيت — Install

```bat
cd "C:\Users\user\Desktop\AI-2026"
python -m pip install -r requirements.txt
```

## 3) التشغيل — Run

**الطريقة الأسهل / easiest:** انقر مرتين على `start.bat`
(يتحقّق من Python، يفتح المتصفح، ويعيد تشغيل السيرفر تلقائياً عند التوقف).

**أو يدوياً / or manually:**

```bat
cd "C:\Users\user\Desktop\AI-2026"
python main.py
```

ثم افتح المتصفح على: <http://127.0.0.1:8000>

### تجهيز Ollama (للتحليل الذكي فقط — Prepare Ollama)

```bat
ollama serve                 REM إن لم يكن يعمل مسبقاً
ollama pull qwen3:4b         REM أو أي موديل تختاره
ollama list                  REM لعرض الموديلات المتاحة
```

ثم من صفحة **الإعدادات** داخل التطبيق: أدخل رابط Ollama (افتراضي `http://127.0.0.1:11434`)،
اختر الموديل من القائمة، ثم «اختبار الاتصال» و«حفظ الإعدادات». قائمة الموديلات تُجلب تلقائياً.

## 4) صيغ الملفات المدعومة — Supported formats

`.xlsx` · `.xlsm` · `.xls` · `.xlsb` · `.csv`

- تُقرأ **كل أوراق** الملف تلقائياً، وتُتخطّى صفوف الترويسة الحكومية وأرقام البنود غير البياناتية.
- يُكتشف **نوع الملف** تلقائياً: مواقف مالية (Financial Position)، سجل مدفوعات (Ledger)، أو مختلط.
- أسماء الأعمدة تُوحَّد (مرادفات عربية/إنجليزية + الأرقام العربية)، والترتيب/اللغة لا يهم.

## 5) ما الذي يكتشفه التدقيق — What the audit detects

تعتمد درجة الخطورة العامة (0–100%) على عدد الملاحظات المصنّفة (عالية/متوسطة/منخفضة).

**بنود المواقف المالية / Financial position items**

| القاعدة | الوصف | المعيار |
|---|---|---|
| `POS_NEAR_CEILING` | بنود استنفدت 90%–100% من المخصص | ISA 315 / ISSAI 1240 |
| `POS_AMENDMENT` | تعديلات على المخصصات المقدرة (مخاطر تجاوز الإدارة) | ISA 240 / ISSAI 1240 |
| `POS_YOY_VARIANCE` | تباينات جوهرية بين السنوات في نسب الإنفاق | ISA 315 |
| `POS_ZERO_BUDGET_PAID` | إنفاق على بنود بمخصص صفري | ISA 315 / ISSAI 1240 |
| `POS_IDLE_ALLOCATION` | مخصصات بلا حركة مالية (بنود خاملة) | ISSAI 1260 |
| `POS_LOW_EXECUTION` | نسبة تنفيذ منخفضة (< 30%) على مخصص كبير | ISSAI 1260 |

**سجلات الحركات المالية / Movement (ledger) records**

| القاعدة | الوصف | المعيار |
|---|---|---|
| `LED_DUPLICATE_PAYMENTS` | مدفوعات مكررة لنفس المستفيد بالمبلغ ذاته | ISA 240 (Fraud) |
| `LED_SPLIT_PURCHASES` | تجزئة مشتريات لتفادي سقف الشراء المباشر | ISA 240 |
| `LED_STAT_OUTLIERS` | مبالغ شاذة إحصائياً | ISA 530 |
| `LED_CONCENTRATION` | تركّز مرتفع للمدفوعات لدى مستفيد واحد | ISSAI 1240 |
| `LED_BELOW_THRESHOLD` | مبالغ أسفل سقف الموافقة مباشرةً (Structuring) | ISA 240 |
| `LED_YEAR_END` | تكدّس الحركات في نهاية السنة المالية | ISSAI 1240 / ISA 560 |
| `LED_ROUND_AMOUNTS` | مبالغ مقرّبة جداً (مضاعفات 1000) | ISSAI 1240 |
| `LED_NEGATIVE_AMOUNTS` | مبالغ سالبة في السجل | ISA 315 / ISA 240 |
| `LED_ZERO_AMOUNTS` | حركات بمبلغ صفر | ISA 315 |
| `LED_WEEKEND_TRANSACTIONS` | حركات في أيام العطل (الجمعة/السبت) | ISA 240 |

## 6) الواجهة — The interface

- **لوحة المعلومات (Dashboard):** نوع الملف + درجة الخطورة + مؤشرات KPI + رسوم بيانية + الملاحظات (قابلة للتوسيع).
- **رفع ملف (Upload):** سحب وإفلات أو اختيار ملف لإجراء التدقيق.
- **التحليل الذكي (AI Analysis):** تقرير تدقيق مهني مُولّد من الملاحظات الفعلية (بثّ مباشر).
- **المحادثة التحليلية (Chat):** أسئلة تفاعلية مرتبطة حصراً بأرقام الملف.
- **تقرير قابل للتصدير (Export):** صفحة HTML مكتفية بذاتها يمكن طباعتها/حفظها PDF.
- الواجهة ثنائية اللغة (عربي/إنجليزي).

## 7) الإعدادات — Configuration (`config.json`)

```json
{
  "ollama_url": "http://127.0.0.1:11434",
  "model": "qwen3:4b",
  "language": "ar",
  "max_ai_records": 40,
  "temperature": 0.2,
  "num_predict": 4096,
  "request_timeout": 300,
  "think": false
}
```

- `max_ai_records`: أقصى عدد ملاحظات تُرسل للموديل.
- `think`: `true` للسماح لموديلات الاستدلال (qwen3…) بالتفكير قبل الإجابة (أدقّ لكن أبطأ).

> **ملاحظة حول الأداء:** بعض الموديلات الكبيرة تعمل على المعالج (CPU) ببطء شديد. للحصول على
> ردود محادثة سريعة، اختر موديلاً أصغر (مثل `qwen3:0.6b`) من صفحة الإعدادات. الموديلات الكبيرة
> (مثل `qwen3:4b`) أدقّ لكنها أبطأ بكثير على الأجهزة بلا GPU قوي.

### العمل دون إنترنت — Offline operation

التطبيق **لا يحتاج إنترنت إطلاقاً**:
- كل ملفات الواجهة (Bootstrap، Font Awesome، Chart.js، خط Cairo) محلية داخل مجلد `static/` — لا CDN.
- التحليل الذكي والمحادثة يعملان عبر Ollama المحلي على `127.0.0.1:11434`.
- عمليات التدقيق وقراءة الملفات وتصدير التقارير تتم بالكامل على جهازك.

> على ويندوز، يُشغّل الخادم حلقة `Selector` (بدل `Proactor` الافتراضية) لأن حلقة `Proactor` قد
> تتوقف عن الاستماع نهائياً بعد قطع اتصال مفاجئ — وهو سبب ظهور رسالة «تعذّر الاتصال بالسيرفر».

## 8) الواجهات البرمجية — API

| الطريقة | المسار | الوظيفة |
|---|---|---|
| `POST` | `/audit/` | رفع ملف (multipart) وإرجاع تقرير JSON كامل |
| `POST` | `/analyze/` | بثّ تقرير ذكي (NDJSON) من ملاحظات التقرير |
| `POST` | `/chat/` | محادثة تحليلية (NDJSON) مرتبطة بأرقام الملف |
| `POST` | `/export/` | تقرير HTML مكتفٍ بذاته |
| `GET`  | `/ollama/models` | الموديلات المتاحة من Ollama |
| `GET`/`POST` | `/settings` | قراءة/حفظ الإعدادات |
| `GET`  | `/health`, `/state` | فحص الصحة وحالة التطبيق |

## 9) بنية المشروع — Project layout

```
AI-2026/
├─ main.py            # خادم FastAPI + نقاط النهاية + تصدير التقرير
├─ audit_engine.py    # محرك التدقيق: القراءة المرنة + القواعد + KPI + الرسوم
├─ index.html         # الواجهة (لوحة معلومات/رفع/تحليل/محادثة/إعدادات) ثنائية اللغة
├─ config.json        # إعدادات Ollama واللغة والموديل
├─ requirements.txt   # الحِزم المطلوبة
├─ start.bat          # مشغّل Windows مع إعادة تشغيل تلقائية
└─ static/            # ملفات ثابتة
```

## 10) استكشاف الأخطاء — Troubleshooting

- **«تعذّر الاتصال بـ Ollama»:** تأكد أن Ollama يعمل (`ollama serve`) والرابط صحيح،
  ثم «اختبار الاتصال» في الإعدادات.
- **لا تظهر موديلات:** نفّذ `ollama pull <model>`.
- **ملف `.xlsb` لا يُقرأ:** تأكد من تثبيت `pyxlsb` (`python -m pip install pyxlsb`).
- **السيرفر يعمل لكن الصفحة لا تستجيب:** أعد تشغيله عبر `start.bat` (يعيد التشغيل تلقائياً).
