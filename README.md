# HK SME Cashflow & Credit Check

A lightweight Python tool for **Hong Kong commercial banking SME credit screening**.
It reads a single Excel workbook (`sample_sme_fin.xlsx`) containing an SME's balance
sheet and income statement, computes the core credit metrics a relationship manager
cares about, assigns a bank-perspective rating (Safe / Watch / Warning) to each metric,
plots a trend chart, and emits a structured first-pass credit memo in Markdown.

> Pure CPU. No GPU/CUDA. No LLM / large-model dependencies.

---

## Project Overview

- **Input:** `sample_sme_fin.xlsx` with two sheets — `BS` (balance sheet) and `IS` (income statement).
- **Output:** `output_chart.png` (charts) and `credit_report.md` (structured screening report).
- **Metrics computed:**
  - Current ratio (流动比率) = Current assets / Current liabilities
  - Quick ratio (速动比率) = (Current assets − Inventory) / Current liabilities
  - Debt ratio (资产负债率) = Total liabilities / Total assets
  - Interest coverage (利息保障倍数) = EBIT / Interest expense
  - Simplified operating cash flow (经营性现金流简易测算) via the indirect method
- **Rating:** each ratio is mapped to 安全 (Safe) / 关注 (Watch) / 预警 (Warning) using
  conventional SME credit-screening thresholds, plus an overall composite rating.

---

## Business Use Case

For **HK commercial banking SME credit screening**, this tool automates the
spread-and-pre-screen step an RM performs before deeper underwriting:

- Quickly derive liquidity, leverage and debt-service coverage ratios from a client's
  financials.
- Flag deteriorating trends (e.g. rising leverage, working-capital build-up, weakening
  quick ratio) as early-warning signals for credit monitoring.
- Produce a reproducible, auditable first-pass credit memo to support the credit
  application narrative and risk-based pricing discussion.

It is designed to **complement my other HK IPO valuation project (on Mac)**:
together they cover the full SME lifecycle — from early-stage **SME financing /
credit screening** all the way to a **Hong Kong IPO valuation** — demonstrating
end-to-end coverage across Commercial Banking and the capital-markets journey.

---

## Features

- Reads multi-period financials (e.g. FY2021–FY2023) from Excel.
- Computes 5 core credit metrics with **fully commented formulas** (see
  `main.py` for line-by-line commentary).
- Assigns 安全 / 关注 / 预警 ratings with explicit, adjustable thresholds.
- Auto-generates `output_chart.png`:
  1. Cash-flow trend (cash earnings vs. working-capital-adjusted operating cash flow)
  2. Liquidity bar chart (current ratio & quick ratio with a 1.0x warning line)
  3. Leverage & coverage chart (debt ratio bars + interest coverage line, dual axis)
- Writes a structured `credit_report.md` (metrics table, thresholds, trend table,
  risk flags, bank-view recommendations).
- Defensive data handling: safe division, NaN-tolerant formatting, period-alignment
  validation between the BS and IS sheets.
- Zero heavyweight dependencies: only `pandas`, `openpyxl`, `matplotlib`.

---

## Input Data Format

`sample_sme_fin.xlsx` must contain two sheets. Column 1 is the item name (项目),
and every subsequent column is one fiscal period. All amounts are in **HKD '000**.

### Sheet `BS` — Balance Sheet (资产负债表)

| 项目 (Item) | FY2021 | FY2022 | FY2023 |
|-------------|--------|--------|--------|
| 货币资金 (Cash) | 4000 | 3800 | 3500 |
| 应收账款 (Accounts receivable) | 6500 | 8200 | 10500 |
| 存货 (Inventory) | 6000 | 7800 | 10000 |
| 预付账款 (Prepayments) | 800 | 1000 | 1200 |
| 流动资产合计 (Total current assets) | 17300 | 20800 | 25200 |
| 固定资产净值 (Net fixed assets) | 7500 | 8600 | 9400 |
| 长期投资 (Long-term investments) | 1000 | 1000 | 1000 |
| 非流动资产合计 (Total non-current assets) | 8500 | 9600 | 10400 |
| 资产总计 (Total assets) | 25800 | 30400 | 35600 |
| 短期借款 (Short-term borrowings) | 4500 | 6000 | 8000 |
| 应付账款 (Accounts payable) | 3800 | 4500 | 5200 |
| 预收账款 (Advances from customers) | 1200 | 1300 | 1400 |
| 其他流动负债 (Other current liabilities) | 1000 | 1100 | 1200 |
| 流动负债合计 (Total current liabilities) | 10500 | 12900 | 15800 |
| 长期借款 (Long-term borrowings) | 4800 | 5300 | 6500 |
| 非流动负债合计 (Total non-current liabilities) | 4800 | 5300 | 6500 |
| 负债合计 (Total liabilities) | 15300 | 18200 | 22300 |
| 实收资本 (Paid-in capital) | 5000 | 5000 | 5000 |
| 未分配利润 (Retained earnings) | 5500 | 7200 | 8300 |
| 所有者权益合计 (Total equity) | 10500 | 12200 | 13300 |
| 负债及所有者权益总计 (Total liabilities & equity) | 25800 | 30400 | 35600 |

### Sheet `IS` — Income Statement (利润表)

| 项目 (Item) | FY2021 | FY2022 | FY2023 |
|-------------|--------|--------|--------|
| 营业收入 (Revenue) | 32000 | 38000 | 45000 |
| 营业成本 (COGS) | 22400 | 26600 | 31500 |
| 毛利 (Gross profit) | 9600 | 11400 | 13500 |
| 销售及管理费用 (SG&A) | 4200 | 5000 | 5900 |
| 利息费用 (Interest expense) | 700 | 850 | 1050 |
| 折旧与摊销 (Depreciation & amortisation) | 1100 | 1300 | 1500 |
| 税前利润 (Profit before tax) | 4700 | 5550 | 6550 |
| 所得税 (Income tax) | 776 | 916 | 1081 |
| 净利润 (Net profit) | 3924 | 4634 | 5469 |

> Note: `折旧与摊销` (D&A) is a disclosed memo item already embedded in COGS/SG&A;
> it is used only for the cash-flow estimation (non-cash add-back).

To build this file yourself, run `python generate_sample_data.py`, or replicate the
two tables above in Excel with the exact item names.

## Metrics, Formulas & Rating Thresholds

All formulas are implemented and commented in `main.py`.

| Metric | Formula | Safe | Watch | Warning |
|--------|---------|------|-------|---------|
| Current ratio | Current assets / Current liabilities | ≥ 2.0x | 1.0x – 2.0x | < 1.0x |
| Quick ratio | (Current assets − Inventory) / Current liabilities | ≥ 1.0x | 0.5x – 1.0x | < 0.5x |
| Debt ratio | Total liabilities / Total assets | ≤ 50% | 50% – 70% | > 70% |
| Interest coverage | EBIT / Interest expense | ≥ 5x | 2x – 5x | < 2x |

**EBIT** = Profit before tax + Interest expense (interest is added back because it is
already deducted in profit before tax; equals operating profit when there are no
material non-operating items).

**Simplified operating cash flow (indirect method):**

```
Cash earnings   = Net profit + Depreciation & amortisation + Interest expense
Adjusted OCF    = Cash earnings − ΔAccounts receivable − ΔInventory + ΔAccounts payable
```

- `Cash earnings` adds back non-cash D&A and interest to obtain a pre-financing
  operating-cash proxy.
- `Adjusted OCF` further adjusts for working-capital movements (receivables/inventory
  build-up consumes cash; payables increase releases cash). The first period has no
  prior-year balance, so its adjusted OCF is reported as `N/A`.

---

## How to Run

```bash
# 1) (Optional) create a virtual environment
python -m venv .venv
# Windows:
.venv\Scripts\activate
# macOS/Linux:
# source .venv/bin/activate

# 2) install dependencies (CPU only)
pip install -r requirements.txt

# 3) generate the sample workbook (if you don't have your own data)
python generate_sample_data.py

# 4) run the credit screening
python main.py
```

Expected outputs:

- `sample_sme_fin.xlsx` — sample financials (from step 3)
- `output_chart.png` — cash-flow trend + solvency bar charts
- `credit_report.md` — structured first-pass credit memo

---

## Output Files

| File | Description |
|------|-------------|
| `output_chart.png` | 3-panel chart: cash-flow trend, liquidity ratios, leverage & coverage |
| `credit_report.md` | Markdown memo: metric table with ratings, threshold reference, multi-period trend table, risk flags, bank-view recommendations |

---

## Verification Checklist (self-check)

Use these checks to confirm the metrics are correct. Expected values are for the
bundled sample data (FY2023, HKD '000).

- [ ] **Current ratio** = 25,200 / 15,800 = **1.59x**
- [ ] **Quick ratio** = (25,200 − 10,000) / 15,800 = **0.96x**
- [ ] **Debt ratio** = 22,300 / 35,600 = **62.6%**
- [ ] **EBIT** = 6,550 + 1,050 = **7,600**
- [ ] **Interest coverage** = 7,600 / 1,050 = **7.24x**
- [ ] **Cash earnings** = 5,469 + 1,500 + 1,050 = **8,019**
- [ ] **Adjusted OCF (FY2023)** = 8,019 − (10,500 − 8,200) − (10,000 − 7,800) + (5,200 − 4,500) = **4,219**
- [ ] **Balance check** — Total assets = Total liabilities + Total equity: 35,600 = 22,300 + 13,300
- [ ] **Sub-total check** — Current assets = 3,500 + 10,500 + 10,000 + 1,200 = 25,200
- [ ] **Chart file** `output_chart.png` exists and is non-empty
- [ ] **Report file** `credit_report.md` exists and contains the rating table

## Interview Talking Points (HSBC Commercial Banking RM)

Use these to explain the project in a CB RM interview, leaning into how a relationship
manager thinks about risk and relationship value.

**What it does (elevator pitch)**
- "It's a lightweight Python tool that takes an SME's balance sheet and income
  statement and produces a first-pass credit screen — liquidity, leverage, debt-service
  coverage and a simplified operating cash-flow estimate — each tagged Safe / Watch /
  Warning, plus a chart and a draft credit memo."

**Why it matters to a Commercial Banking RM**
- "RMs pre-screen borrowers before formal underwriting. This automates the spread and
  flags deterioration early, so I can have a sharper credit conversation with the client
  and with Credit Risk — e.g. 'your quick ratio dropped below 1.0x because inventory
  built up faster than payables'."
- "It supports risk-based pricing and early-warning monitoring, which is core to
  HSBC's SME lending discipline: know the client's cash conversion cycle, not just the
  collateral."

**How I read the ratios (bank lens)**
- Current ratio / quick ratio → can the client meet short-term obligations; quick ratio
  strips out inventory that may be slow to liquidate.
- Debt ratio → overall leverage; a rising trend is a red flag even if absolute level is
  still acceptable.
- Interest coverage (EBIT / interest) → how many times operating earnings cover interest;
  the key debt-service buffer.
- Operating cash flow → whether profit is converting to cash; positive but
  working-capital-dragged OCF signals growth straining liquidity.

**Why it complements my HK IPO valuation project (on Mac)**
- "Together they cover the SME lifecycle end-to-end: this tool screens and supports SME
  financing at the Commercial Banking stage, while my other project values the company
  at the HK IPO stage. That shows I understand both credit and capital markets — from
  lending to listing."

**Technical & professional credibility**
- "It's deliberately lightweight — pandas/openpyxl/matplotlib, pure CPU, no GPU and no
  LLM dependencies — so it's easy to run, audit and explain to a colleague or a credit
  officer. Every formula is commented and reproducible."

**Tie-back to HSBC**
- "This mirrors HSBC's focus on International SME clients and trade — the working-capital
  adjustments (receivables, inventory, payables) are exactly the lines an RM would look
  at when structuring trade finance, invoice financing or a revolving facility."

---

## Project Structure

```
hk-sme-cashflow-credit-check/
├── main.py                 # main screening logic (read → compute → rate → plot → report)
├── generate_sample_data.py # creates sample_sme_fin.xlsx (structure + mock data)
├── sample_sme_fin.xlsx     # sample input workbook (generated)
├── output_chart.png        # generated chart
├── credit_report.md        # generated screening memo
├── requirements.txt        # CPU-only dependencies
└── README.md               # this file
```

---

## Disclaimer

The bundled data is synthetic, for demonstration and self-testing only, and does not
represent any real client. Rating thresholds are generic SME credit-screening
guidelines and should be calibrated to your bank's policy and the client's industry.
This tool is a decision-support aid, not a substitute for formal credit underwriting.


