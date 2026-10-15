# لوحة مؤشرات الأداء المالي — دليل مختصر
# Financial KPI Dashboard — Concise Guide

---

## العربية

### 1) المصدر (من أين تأتي الأرقام؟)
- **ملف العمل:** `Finance_KPI_Master_Framework_F01-F10_Odoo_Dashboard_v10_DASHBOARDS_COMPLETE.xlsx` (82 ورقة).
- **الشركة:** DesignTech Engineering Consultants — **الفترة:** السنة المالية 2025 — **كما في:** 2025-12-31 — **العملة:** SAR.
- **الأوراق الأساسية الست:** دفتر الأستاذ العام (General Ledger، 249,377 حركة)، الميزانية العمومية (Balance Sheet)، الأرباح والخسائر (P&L)، التدفقات النقدية (Cash Flow)، دليل الحسابات (Chart of Accounts، 443)، الحسابات التحليلية (Analytic Accounts، 164).
- **أوراق الإطار:** `KPI Master` (60 مؤشر)، `Data Mapping` (76 مفتاح قياس)، `KPI Calculation`، بالإضافة إلى `Executive` و`Finance` و10 لوحات أهداف.
- **6 ملفات تشغيلية إضافية (مصدر التحليل التشغيلي):** `DT2025_Customer_Contract_Data.xlsx` (العملاء + العقود)، `DT2025_VAT_Tax_Compliance.xlsx` (الإقرار الضريبي)، `DT2025_Bank_Reconciliations.xlsx` (المطابقات البنكية)، `DT2025_Procurement_Data_Purchase Order1.xlsx` (المشتريات)، `DT2025_Cost_Center_Master_Data.xlsx` (مراكز التكلفة)، `DT2025_Loans_Repayment_Schedule.xlsx` (القروض).
- يُبنى ملف `data.js` من هذه الأوراق عبر `build_data.js` + `pipeline.js` + `ops_build.js`، ثم تعرضه الواجهة في `app.js`.
- كل مفتاح في `Data Mapping` يُطابق من ورقته المصدر بـ**آخر صف** يحمل نفس العنوان (مطابقة `LOOKUP(2,1/...)` في ملف العمل).

### 2) القوائم (شريط التنقّل الجانبي)
| القائمة | ماذا تعرض | المصدر |
| --- | --- | --- |
| **اللوحة التنفيذية** `exec` | درجة إجمالية، عدد الأهداف الحرجة/التي تحتاج متابعة، مؤشرات محسوبة، جدول الأهداف، مخططان (حالة المؤشرات + درجات الأهداف) | ملخّص محسوب من الإطار |
| **لوحة المالية** `fin` | نفس الملخّص + جدول الأهداف ومخطط أعمدة لدرجات الأهداف | ملخّص محسوب من الإطار |
| **سجل المؤشرات** `master` | كل المؤشرات الـ60 مع الصيغة والقيمة والهدف والدرجة والحالة (فلترة بالهدف/الحالة) | `KPI Master` + `KPI Calculation` |
| **خريطة البيانات** `mapping` | الـ76 مفتاحًا: المصدر، الوصف، القيمة، الجاهزية | `Data Mapping` |
| **الهرم** (هدف ← مؤشر) | عند النقر على هدف تظهر مؤشراته؛ وعند النقر على مؤشر تظهر صفحته التفصيلية (التعريف، التسلسل، المدرج التكراري للدرجة) | محسوب |
| **التحليل التشغيلي** `ops` | 6 تبويبات: الذمم/التحصيل، المبيعات، المشتريات، الضريبة، مراكز التكلفة، الرقابة والخزينة — انظر القسم 8 | 6 ملفات Odoo التشغيلية |
| **المستكشف التفصيلي** | تبويبات: **كل الحركات** (249,377)، **المشاريع** (163)، **دليل الحسابات** (442)، **العملاء/الموردون** (899) — انظر القسم 4 | الأوراق الأساسية |

### 3) القيم الرئيسية ومصدر كل قيمة
| القيمة | المبلغ (SAR) | المصدر |
| --- | --- | --- |
| الإيرادات | 172,031,968.76 | قائمة الأرباح والخسائر |
| إجمالي الدخل (الإيرادات + إيرادات أخرى) | 174,080,152.99 | قائمة الأرباح والخسائر |
| تكلفة الإيرادات | 104,351,012.85 | قائمة الأرباح والخسائر |
| الربح الإجمالي | 67,680,955.91 | قائمة الأرباح والخسائر |
| المصروفات التشغيلية | 7,495,424.18 | قائمة الأرباح والخسائر |
| صافي الربح | 62,233,715.96 | قائمة الأرباح والخسائر |
| إجمالي الأصول | 73,379,315.05 | الميزانية العمومية |
| الأصول المتداولة | 58,797,720.30 | الميزانية العمومية |
| النقدية | 73,715,524.08 | الميزانية العمومية |
| إجمالي الالتزامات | 10,990,826.19 | الميزانية العمومية |
| حقوق الملكية | 62,388,488.86 | الميزانية العمومية |
| النقدية الافتتاحية / الختامية | 29,226,629.49 / 72,096,315.17 | قائمة التدفقات النقدية |

**نسب محسوبة:** النسبة الجارية 26.87 · رأس المال العامل 56,609,861.89 · نسبة النقدية 33.69 · الدين/الملكية 0.18 · ROA 84.81 · ROE 99.75 · هامش الربح الإجمالي 39.34% · هامش الربح الصافي 35.75%.

### 4) المستكشف التفصيلي (6 تبويبات / 4 شاشات)
- **كل الحركات:** كل سطور دفتر الأستاذ (الحركة، الحساب، الوصف، الشريك، مدين، دائن) مع ترقيم الصفحات (150/صفحة) وفلاتر: **الفترة** + **الحساب** + **العملاء**.
- **المشاريع:** 163 مشروعًا (الاسم، المرجع، العميل، الخطة، مدين/دائن/الرصيد) + مخطط أعلى 12 مشروعًا بالرصيد. فلتر: العميل.
- **دليل الحسابات:** 442 حسابًا (بدون حسابات "Unmapped" التجميعية ذات الرصيد الصفري) مع عدد الحركات لكل حساب + مخطط المدين/الدائن حسب النوع. فلتر: النوع (17 نوعًا).
- **العملاء/الموردون:** تجميع الحركات حسب الشريك (عدد الحركات، مدين، دائن، صافي) + مخطط أعلى 12. فلتر: الفترة.
- **النقر على أي صف (مشروع / حساب / عميل):** يفتح **بوباب بالتفاصيل** — ملخص (مدين/دائن/رصيد/عدد الحركات)، رسم بياني، وجدول حركات خاصة به **فقط**. يحتوي البوباب زر «فتح الحركات في المستكشف». صف الهدف/المؤشر → صفحته. يُغلق البوباب بالزر ✕ أو بالنقر خارجه أو بمفتاح Esc.

### 5) الرسوم البيانية (تعمل مع البيانات الكبيرة)
- **اللوحة التنفيذية/المالية:** مخطط دائري لحالة المؤشرات + مخطط أعمدة لدرجات الأهداف (0–100).
- **كل الحركات:** أعمدة شهرية (مدين مقابل دائن) حسب الفلتر الحالي.
- **المشاريع:** أعلى 12 مشروعًا حسب |الرصيد| (أخضر موجب / أحمر سالب).
- **دليل الحسابات:** أعمدة مكدّسة (مدين/دائن) حسب نوع الحساب.
- **العملاء:** أعلى 12 شريكًا حسب حجم الحركة.
- جميع المخططات مبنية بـ ECharts وتُعاد رسمها عند تغيير الفلتر أو الحجم أو المظهر.

### 6) الفلاتر
- **الفترة:** «الكل» + 12 شهرًا (2025-01 … 2025-12). تظهر في: الحركات، العملاء.
- **الحساب:** «الكل» + 442 حسابًا برمزها واسمها (الحركات فقط).
- **العملاء:** «الكل» + كل الشركاء + «غير محدد» (الحركات والمشاريع).
- **النوع:** «الكل» + 17 نوعًا (دليل الحسابات فقط).
- **الهدف/الحالة:** (سجل المؤشرات فقط). زر «تصفير» يعيد الضبط، والرقائق تُظهر الفلاتر النشطة.

### 7) ملاحظات الدقة
- إجمالي مدين = إجمالي دائن = **4,402,440,355.86** (مجموع الـ442 حسابًا، مطابق لصف الإجمالي في دفتر الأستاذ)، والرصيد = 0.
- **لا استبدال للصفر:** أي مؤشر ينقصه مدخل يظهر «بيانات مطلوبة» بدل صفر.
- **الدرجة (0–100):** داخل النطاق = 100؛ أقل من الحد = القيمة/الحد×100؛ أعلى من الحد = الحد/القيمة×100.
- **الحالة:** لا قيمة → *بيانات مطلوبة*؛ لا هدف رقمي → *بانتظار الهدف*؛ ≥100 *على الهدف*؛ ≥80 *تحتاج متابعة*؛ وإلا *حرجة*.
- القيود: **25 مؤشرًا من 60 «بيانات مطلوبة»** (لا تتوفر مدخلاتها في أي من الملفات — بيانات موازنة/توقع/تدقيق/سياسات/سجل مخاطر/أوقات الإغلاق)، ولا يمكن حساب الاتجاهات (تصدير لفترة واحدة).
- **حالة الحساب بعد الإثراء:** 60 مؤشرًا · **35 محسوبًا** · **16 مُقيَّمًا** · الدرجة الإجمالية **52.8/100** · التوزيع: حرجة 9 · على الهدف 7 · بانتظار الهدف 19 · بيانات مطلوبة 25.
- **مؤشرات مُشتقَّقة من الملفات المتوفرة** (قواعد معلنة، وتظهر في مدخلات كل مؤشر بمصدرها الفعلي):
  - **دورة التحويل النقدي** (F04-K03) = DSO + أيام المخزون − DPO = 31.89 يومًا.
  - **التدفق النقدي الحر** (F03-K03) = التدفق التشغيلي (64,642.66) − النفقات الرأسمالية (1,191,744.70 من حركة الأصول الثابتة في دفتر الأستاذ) = −1,127,102.04.
  - **تحويل النقد للمشاريع** (F07-K05) = المحصّل / المفوتر = 199,818,066.73 / 199,933,128.57 = 99.94%.
  - **التكلفة لكل وحدة إيراد** (F05-K06) = المصروفات / عدد وحدات الإيراد (89 مشروعًا) = 84,218.25.
  - **تغطية التقارير الإدارية** (F08-K04) = المؤشرات ذات القيمة / المؤشرات المطلوبة = 35/60 = 58.3%.
  - **تغطية نسب البيانات** (F08-K06) = المدخلات ذات القيمة / إجمالي المدخلات المعيَّنة = 40/76 = 52.6%.
  - **معدل الحركات غير المعتادة** (F09-K05) = معاملات تتجاوز المتوسط + 3 انحرافات معيارية / إجمالي الحركات = 811/249,377 = 0.33%.
  - **التعرض المالي للمخاطر** (F10-K01) = الذمم المتأخرة أكثر من 90 يومًا = 115,061.84.
  - **التعرض التعاقدي** (F10-K05) = المتبقي (الذمم المستحقة) / إجمالي المفوتر = 0.06%.

### 8) التحليل التشغيلي (`ops`) — من ملفات Odoo الستة
يُبنى بواسطة `ops_build.js` (يقرأ الملفات الستة من مجلد المشروع). التبويبات:
| التبويب | المحتوى | المصدر |
| --- | --- | --- |
| **الذمم والتحصيل** `ar` | مفوتر 199,933,128.57 · محصّل 199,821,465.13 · مستحق/متأخر 115,061.84 · تقادم أعمار الديون · أعلى العملاء (3,973 فاتورة) | Customer Contract Data |
| **المبيعات** `sales` | 53,097 أمر بيع بإجمالي 245,897,868.15 · مخطط شهري · أعلى المندوبين والعملاء | Sales Orders |
| **المشتريات** `proc` | 312 أمر شراء · مخطط شهري · أعلى الموردين · جدول الأوامر | Procurement Data |
| **الضريبة** `vat` | ضريبة المبيعات (مخرجات) 25,724,594.48 · ضريبة المشتريات (مدخلات) 5,388,702.20 · **صافي المستحق 20,335,892.28** · 9,252 سطرًا · مخطط شهري | VAT & Tax Compliance |
| **مراكز التكلفة** `cc` | 89 مركزًا · إيراد 176,979,336.68 · تكلفة 116,333,628.97 · 31 مشروعًا خاسرًا | Cost Center Master Data |
| **الرقابة والخزينة** `ctrl` | مطابقة بنكية 100% (4,764/4,764) · تدفقات داخل 521,158,727.92 / خارج 480,195,684.39 · التسهيلات والقروض | Bank Reconciliations + Loans |

- ملاحظة: قسم «التحليل التشغيلي» يعرض **السنة المالية كاملة** (2025) دون فلتر فترة (الفلترة متاحة في المستكشف).

### 9) اللغة والتسمية
- الواجهة تدعم **العربية والإنجليزية** بالكامل؛ أسماء **الأهداف (F01–F10)** و**المؤشرات الستين** مُعرَّبة وتظهر عربيةً في الوضع العربي وإنجليزيةً في الوضع الإنجليزي.
- أسماء الأطراف/الحسابات/المشاريع (بيانات المصدر) تبقى بحروفها الأصلية في الوضعين.

---

## English

### 1) Source
- **Workbook:** `Finance_KPI_Master_Framework_F01-F10_Odoo_Dashboard_v10_DASHBOARDS_COMPLETE.xlsx` (82 sheets).
- **Company:** DesignTech Engineering Consultants — **Period:** FY 2025 — **As of:** 2025-12-31 — **Currency:** SAR.
- **Six base sheets:** General Ledger (249,377 rows), Balance Sheet, Profit & Loss, Cash Flow, Chart of Accounts (443), Analytic Accounts (164).
- **Framework sheets:** `KPI Master` (60 KPIs), `Data Mapping` (76 metric keys), `KPI Calculation`, plus `Executive`/`Finance` and 10 objective dashboards.
- **Six extra operational files (source of Operational Analytics):** `DT2025_Customer_Contract_Data.xlsx` (customers + contracts), `DT2025_VAT_Tax_Compliance.xlsx`, `DT2025_Bank_Reconciliations.xlsx`, `DT2025_Procurement_Data_Purchase Order1.xlsx`, `DT2025_Cost_Center_Master_Data.xlsx`, `DT2025_Loans_Repayment_Schedule.xlsx`.
- `data.js` is built from these sheets by `build_data.js` + `pipeline.js` + `ops_build.js` and rendered by `app.js`.
- Each `Data Mapping` metric resolves from its source sheet by **label, taking the LAST matching row** (the workbook uses `LOOKUP(2,1/...)` = last match).

### 2) Menus (sidebar)
| Menu | Shows | Source |
| --- | --- | --- |
| **Executive Dashboard** `exec` | Overall score, critical/attention objectives, computed count, objective table, status pie + objective score bar | Computed summary |
| **Finance Dashboard** `fin` | Same summary + objective table + objective score bar | Computed summary |
| **KPI Master** `master` | All 60 KPIs with formula, actual, target, score, status (filter by objective/status) | `KPI Master` + `KPI Calculation` |
| **Data Mapping** `mapping` | The 76 keys: source, label, value, readiness | `Data Mapping` |
| **Hierarchy** (objective → KPI) | Click an objective to see its KPIs; click a KPI for its detail page (definition, lineage, score gauge) | Computed |
| **Operational Analytics** `ops` | 6 tabs: AR/Collections, Sales, Procurement, VAT, Cost Centers, Control & Treasury — see §8 | 6 Odoo operational files |
| **Detailed Explorer** | Tabs: **All Transactions** (249,377), **Projects** (163), **Chart of Accounts** (442), **Customers/Vendors** (899) — see §4 | Base sheets |

### 3) Key values and their sources
| Value | Amount (SAR) | Source |
| --- | --- | --- |
| Revenue | 172,031,968.76 | P&L |
| Total Income (revenue + other income) | 174,080,152.99 | P&L |
| Cost of Revenue | 104,351,012.85 | P&L |
| Gross Profit | 67,680,955.91 | P&L |
| Operating Expenses | 7,495,424.18 | P&L |
| Net Profit | 62,233,715.96 | P&L |
| Total Assets | 73,379,315.05 | Balance Sheet |
| Current Assets | 58,797,720.30 | Balance Sheet |
| Cash | 73,715,524.08 | Balance Sheet |
| Total Liabilities | 10,990,826.19 | Balance Sheet |
| Equity | 62,388,488.86 | Balance Sheet |
| Opening / Closing Cash | 29,226,629.49 / 72,096,315.17 | Cash Flow |

**Computed ratios:** Current 26.87 · Working Capital 56,609,861.89 · Cash Ratio 33.69 · Debt/Equity 0.18 · ROA 84.81 · ROE 99.75 · Gross Margin 39.34% · Net Margin 35.75%.

### 4) Detailed Explorer (4 screens / 6 tabs)
- **All Transactions:** every GL line (date, account, description, partner, debit, credit) with paging (150/page) and filters: **period** + **account** + **customer**.
- **Projects:** 163 projects (name, reference, customer, plan, debit/credit/balance) + top-12-by-balance bar. Filter: customer.
- **Chart of Accounts:** 442 accounts (excluding zero-balance "Unmapped" roll-ups) with movement count per account + debit/credit-by-type bar. Filter: type (17).
- **Customers/Vendors:** transactions grouped by partner (count, debit, credit, net) + top-12 bar. Filter: period.
- **Click any row (project / account / customer):** opens a **details popup** — summary (debit/credit/balance/movement count), a chart, and a movements table for **that item only**, with an "Open movements in explorer" button. Objective/KPI row → its page. Close the popup via ✕, clicking outside, or Esc.

### 5) Charts (built for large data)
- **Executive/Finance:** status pie + objective score bar (0–100).
- **All Transactions:** monthly bars (debit vs credit) for the current filter.
- **Projects:** top 12 by |balance| (green positive / red negative).
- **Chart of Accounts:** stacked bars (debit/credit) by account type.
- **Customers:** top 12 partners by activity.
- All charts use ECharts and redraw on filter, resize, or theme/lang change.

### 6) Filters
- **Period:** All + 12 months (2025-01 … 2025-12). Shown in: Transactions, Customers.
- **Account:** All + 442 accounts by code and name (Transactions only).
- **Customers:** All + every partner + "Unassigned" (Transactions and Projects).
- **Type:** All + 17 types (Chart of Accounts only).
- **Objective/Status:** (KPI Master only). A "Reset" button clears filters; chips show active filters.

### 7) Accuracy notes
- Total debit = total credit = **4,402,440,355.86** (sum of the 442 accounts, matching the GL Total row), balance = 0.
- **No zero substitution:** any KPI missing an input shows "Data Required" instead of zero.
- **Score (0–100):** inside band = 100; below low = actual/low×100; above high = high/actual×100.
- **Status:** missing actual → *Data Required*; no numeric target → *Target Pending*; ≥100 *On Target*; ≥80 *Attention*; else *Critical*.
- Constraints: **25 of 60 KPIs are "Data Required"** (inputs absent from every file — budget/forecast/audit/policy/risk-register/closing-time), and trends are not possible (single-period export).
- **Framework status after enrichment:** 60 KPIs · **35 computed** · **16 scored** · overall **52.8/100** · distribution: Critical 9 · On Target 7 · Target Pending 19 · Data Required 25.
- **Metrics derived from the available files** (rules stated; each KPI shows its real source in its input list):
  - **Cash Conversion Cycle** (F04-K03) = DSO + DIO − DPO = 31.89 days.
  - **Free Cash Flow** (F03-K03) = Operating CF (64,642.66) − Capex (1,191,744.70, from Fixed-Asset movement in the General Ledger) = −1,127,102.04.
  - **Project Cash Conversion** (F07-K05) = Collected / Invoiced = 199,818,066.73 / 199,933,128.57 = 99.94%.
  - **Cost per Revenue Unit** (F05-K06) = Expenses / Revenue Units (89 projects) = 84,218.25.
  - **Management Reporting Coverage** (F08-K04) = KPIs with a value / required KPIs = 35/60 = 58.3%.
  - **Data Lineage Coverage** (F08-K06) = resolved inputs / total mapped inputs = 40/76 = 52.6%.
  - **Unusual Transaction Rate** (F09-K05) = transactions beyond mean + 3σ / total transactions = 811/249,377 = 0.33%.
  - **Risk Exposure** (F10-K01) = receivables overdue > 90 days = 115,061.84.
  - **Contractual Exposure** (F10-K05) = outstanding (due) / total invoiced = 0.06%.

### 8) Operational Analytics (`ops`) — from the six Odoo files
Built by `ops_build.js` (reads the six files from the project folder). Tabs:
| Tab | Content | Source |
| --- | --- | --- |
| **AR & Collections** `ar` | Invoiced 199,933,128.57 · collected 199,821,465.13 · outstanding/overdue 115,061.84 · aging · top customers (3,973 invoices) | Customer Contract Data |
| **Sales** `sales` | 53,097 orders totalling 245,897,868.15 · monthly chart · top salespeople & customers | Sales Orders |
| **Procurement** `proc` | 312 purchase orders · monthly chart · top vendors · orders table | Procurement Data |
| **VAT** `vat` | Output VAT 25,724,594.48 · input VAT 5,388,702.20 · **net payable 20,335,892.28** · 9,252 lines · monthly chart | VAT & Tax Compliance |
| **Cost Centers** `cc` | 89 centers · revenue 176,979,336.68 · cost 116,333,628.97 · 31 loss-making projects | Cost Center Master Data |
| **Control & Treasury** `ctrl` | 100% bank reconciliation (4,764/4,764) · inflows 521,158,727.92 / outflows 480,195,684.39 · facilities & loans | Bank Reconciliations + Loans |

- Note: the Operational Analytics section shows the **full FY 2025** (no period filter; period filtering lives in the Explorer).

### 9) Language & naming
- The UI is fully **bilingual (Arabic/English)**; the **objective (F01–F10)** and **all 60 KPI** names are translated and render Arabic in Arabic mode and English in English mode.
- Source-data names (parties, accounts, projects) stay in their original script in both modes.
